#!/usr/bin/env python3
"""Run a captured, variable-aware MCP tool sequence over stdio JSON-RPC."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import threading
from pathlib import Path
from typing import Any


def send(process: subprocess.Popen[str], message: dict[str, Any]) -> None:
    assert process.stdin is not None
    process.stdin.write(json.dumps(message, separators=(",", ":")) + "\n")
    process.stdin.flush()


def receive(process: subprocess.Popen[str], request_id: int) -> dict[str, Any]:
    assert process.stdout is not None
    for line in process.stdout:
        if not line.strip():
            continue
        message = json.loads(line)
        if message.get("id") == request_id:
            return message
    raise RuntimeError(f"server closed before response {request_id}")


def drain_stderr(process: subprocess.Popen[str], quiet: bool) -> None:
    assert process.stderr is not None
    for line in process.stderr:
        if not quiet:
            print(line.rstrip(), file=sys.stderr)


def substitute(value: Any, variables: dict[str, Any]) -> Any:
    if isinstance(value, str) and value.startswith("$"):
        name = value[1:]
        if name not in variables:
            raise KeyError(f"workflow variable is not captured yet: {name}")
        return variables[name]
    if isinstance(value, list):
        return [substitute(item, variables) for item in value]
    if isinstance(value, dict):
        return {key: substitute(item, variables) for key, item in value.items()}
    return value


def decoded_tool_result(message: dict[str, Any]) -> Any:
    result = message.get("result", {})
    content = result.get("content", [])
    if content and content[0].get("type") == "text":
        value = content[0].get("text", "")
        try:
            return json.loads(value)
        except json.JSONDecodeError:
            return value
    return result.get("structuredContent", result)


def capture_path(value: Any, path: str) -> Any:
    current = value
    for part in path.split("."):
        if part == "length" and isinstance(current, (dict, list, str)):
            current = len(current)
        else:
            current = current[int(part)] if isinstance(current, list) else current[part]
    return current


def result_failed(result: Any) -> bool:
    return isinstance(result, dict) and (
        result.get("success") is False
        or result.get("status") in {"failed", "partial"}
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--calls", required=True, type=Path, help="workflow JSON array")
    parser.add_argument("--trace", required=True, type=Path, help="full JSON trace output")
    parser.add_argument("--quiet-server", action="store_true", help="suppress server stderr")
    parser.add_argument("command", nargs=argparse.REMAINDER, help="-- SERVER [ARGS ...]")
    args = parser.parse_args()
    if args.command and args.command[0] == "--":
        args.command = args.command[1:]
    if not args.command:
        parser.error("a server command is required after --")
    return args


def main() -> int:
    args = parse_args()
    calls = json.loads(args.calls.read_text(encoding="utf-8"))
    if not isinstance(calls, list):
        raise ValueError("workflow root must be a JSON array")

    process = subprocess.Popen(
        args.command,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=1,
    )
    threading.Thread(
        target=drain_stderr,
        args=(process, args.quiet_server),
        daemon=True,
    ).start()

    responses: list[dict[str, Any]] = []
    variables: dict[str, Any] = {}
    initialized: dict[str, Any] | None = None
    request_id = 1
    try:
        send(
            process,
            {
                "jsonrpc": "2.0",
                "id": request_id,
                "method": "initialize",
                "params": {
                    "protocolVersion": "2024-11-05",
                    "capabilities": {},
                    "clientInfo": {"name": "docx-smoke-probe", "version": "1"},
                },
            },
        )
        initialized = receive(process, request_id)
        send(process, {"jsonrpc": "2.0", "method": "notifications/initialized"})

        for call in calls:
            request_id += 1
            call_id = call.get("id", call["name"])
            print(f"[smoke] {call_id}", file=sys.stderr)
            arguments = substitute(call.get("arguments", {}), variables)
            send(
                process,
                {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "method": "tools/call",
                    "params": {"name": call["name"], "arguments": arguments},
                },
            )
            message = receive(process, request_id)
            decoded = decoded_tool_result(message)
            entry = {
                "id": call.get("id"),
                "tool": call["name"],
                "arguments": arguments,
                "isError": bool(message.get("error"))
                or bool(message.get("result", {}).get("isError", False)),
                "result": decoded,
            }
            assertions = []
            for path, expected in call.get("expect", {}).items():
                actual = capture_path(decoded, path)
                assertions.append(
                    {
                        "path": path,
                        "expected": expected,
                        "actual": actual,
                        "passed": actual == expected,
                    }
                )
            if assertions:
                entry["assertions"] = assertions
            responses.append(entry)
            for name, path in call.get("capture", {}).items():
                variables[name] = capture_path(decoded, path)

        payload = {
            "initialize": initialized,
            "responses": responses,
            "variables": variables,
        }
        args.trace.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")

        transport_errors = sum(bool(entry["isError"]) for entry in responses)
        failed_results = sum(result_failed(entry["result"]) for entry in responses)
        assertion_count = sum(len(entry.get("assertions", [])) for entry in responses)
        failed_assertions = sum(
            not assertion["passed"]
            for entry in responses
            for assertion in entry.get("assertions", [])
        )
        summary = {
            "calls": len(responses),
            "transportErrors": transport_errors,
            "failedResults": failed_results,
            "assertions": assertion_count,
            "failedAssertions": failed_assertions,
            "trace": str(args.trace),
        }
        print(json.dumps(summary, indent=2))
        return 1 if transport_errors or failed_results or failed_assertions else 0
    finally:
        if process.poll() is None:
            request_id += 1
            try:
                send(process, {"jsonrpc": "2.0", "id": request_id, "method": "shutdown"})
                receive(process, request_id)
            except (BrokenPipeError, RuntimeError):
                pass
            process.terminate()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()


if __name__ == "__main__":
    raise SystemExit(main())
