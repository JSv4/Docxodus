"""NDJSON stdio transport to the ``docxodus-pyhost`` subprocess.

One host process per Python process; many ``DocxSession`` handles inside it.
The session-handle pool lives on the .NET side (`SessionRegistry`) — this
module is purely the byte-pipe.

Threading model:
- ``threading.Lock`` serializes one request/response round-trip at a time so
  multiple threads sharing a session never interleave NDJSON lines.
- Request IDs are monotonic per process; the reader matches replies to
  callers by id. (Lock simplification means we don't strictly need id
  matching today, but the wire protocol allows it so a future async client
  can lift the lock without changing the host.)
- A daemon thread drains stderr into Python's ``logging`` so host
  diagnostics aren't silently lost.

Lifecycle:
- Lazy spawn on first ``call()``.
- Process re-spawn if the previous one died (e.g. user killed it externally).
- ``atexit`` hook sends ``shutdown``, waits up to 2 s, then terminates, then
  kills. Best-effort: an interpreter that goes down hard (segfault) leaves
  the host to notice its parent's stdin close and exit on its own.
"""

from __future__ import annotations

import atexit
import json
import logging
import subprocess
import sys
import threading
from typing import Any

from ._host_locator import find_host
from .errors import DocxodusTransportError

_log = logging.getLogger("docx_scalpel.host")


class _Transport:
    """Singleton subprocess + request/response loop."""

    _instance: "_Transport | None" = None
    _instance_lock = threading.Lock()

    def __init__(self, proc: subprocess.Popen[bytes]) -> None:
        self._proc = proc
        self._lock = threading.Lock()
        self._next_id = 0
        self._closed = False
        self._stderr_thread = threading.Thread(
            target=self._drain_stderr, name="docx-scalpel-host-stderr", daemon=True
        )
        self._stderr_thread.start()

    # -- public surface ---------------------------------------------------

    @classmethod
    def get(cls) -> "_Transport":
        """Return the live transport, spawning the host on first use or after a crash."""
        with cls._instance_lock:
            inst = cls._instance
            if inst is None or inst._proc.poll() is not None:
                inst = cls._spawn()
                cls._instance = inst
                atexit.register(inst.shutdown)
            return inst

    def call(self, op: str, args: dict[str, Any] | None = None) -> Any:
        """Send one request, return its ``result``.

        Raises :class:`DocxodusTransportError` on transport-level failures
        (host died, malformed reply, ``ok: false`` envelope from the host).
        Business-level failures inside ``ok: true`` envelopes are passed
        through as their normal Python values — typically a dict the caller
        decodes into an :class:`~docx_scalpel.types.EditResult`.
        """
        payload = json.dumps(
            {"id": -1, "op": op, "args": args or {}},
            separators=(",", ":"),
            ensure_ascii=False,
        )

        with self._lock:
            if self._closed or self._proc.poll() is not None:
                raise DocxodusTransportError(
                    f"host process is no longer running (exit code {self._proc.returncode})"
                )

            self._next_id += 1
            request_id = self._next_id
            # Splice the real id in (avoids re-serializing).
            framed = payload.replace('"id":-1', f'"id":{request_id}', 1) + "\n"

            try:
                assert self._proc.stdin is not None
                self._proc.stdin.write(framed.encode("utf-8"))
                self._proc.stdin.flush()
            except (BrokenPipeError, OSError) as ex:
                raise DocxodusTransportError(
                    f"failed writing request {op!r} to host: {ex}"
                ) from ex

            assert self._proc.stdout is not None
            line = self._proc.stdout.readline()
            if not line:
                self._closed = True
                raise DocxodusTransportError(
                    f"host exited unexpectedly while awaiting reply to {op!r} "
                    f"(exit code {self._proc.poll()})"
                )

            try:
                reply = json.loads(line)
            except json.JSONDecodeError as ex:
                raise DocxodusTransportError(
                    f"malformed reply from host (op={op!r}): {ex}; raw={line!r}"
                ) from ex

            if reply.get("id") != request_id:
                raise DocxodusTransportError(
                    f"reply id mismatch: expected {request_id}, got {reply.get('id')!r}"
                )

            if not reply.get("ok"):
                err = reply.get("error") or {}
                raise DocxodusTransportError(
                    err.get("message", "host returned ok=false with no message"),
                    code=err.get("code"),
                    trace=err.get("trace"),
                )

            return reply.get("result")

    def shutdown(self, *, timeout: float = 2.0) -> None:
        """Send ``shutdown``, then terminate, then kill. Idempotent."""
        with self._lock:
            if self._closed:
                return
            self._closed = True
            proc = self._proc

            if proc.poll() is None and proc.stdin is not None:
                try:
                    proc.stdin.write(b'{"id":0,"op":"shutdown","args":{}}\n')
                    proc.stdin.flush()
                except (BrokenPipeError, OSError):
                    pass
                try:
                    proc.stdin.close()
                except OSError:
                    pass

            try:
                proc.wait(timeout=timeout)
            except subprocess.TimeoutExpired:
                proc.terminate()
                try:
                    proc.wait(timeout=timeout)
                except subprocess.TimeoutExpired:
                    proc.kill()
                    try:
                        proc.wait(timeout=timeout)
                    except subprocess.TimeoutExpired:
                        _log.error("docxodus-pyhost did not exit after kill()")

    # -- internals --------------------------------------------------------

    @classmethod
    def _spawn(cls) -> "_Transport":
        host_path = find_host()
        _log.debug("spawning docxodus-pyhost: %s", host_path)
        proc = subprocess.Popen(
            [str(host_path)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            bufsize=0,
            close_fds=(sys.platform != "win32"),
        )
        return cls(proc)

    def _drain_stderr(self) -> None:
        assert self._proc.stderr is not None
        try:
            for raw in iter(self._proc.stderr.readline, b""):
                line = raw.decode("utf-8", errors="replace").rstrip()
                if line:
                    _log.info("[host] %s", line)
        except Exception:  # noqa: BLE001 — daemon thread should never propagate
            _log.exception("stderr drain thread crashed")


def call(op: str, args: dict[str, Any] | None = None) -> Any:
    """Module-level convenience wrapper around the singleton transport."""
    return _Transport.get().call(op, args)


def shutdown_host() -> None:
    """Best-effort shutdown of the singleton host. Mostly for tests."""
    with _Transport._instance_lock:
        inst = _Transport._instance
        if inst is not None:
            inst.shutdown()
            _Transport._instance = None
