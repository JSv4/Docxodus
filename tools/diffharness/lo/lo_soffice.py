"""Start and stop an isolated headless soffice without holding the caller's pipes open.

`soffice` is a wrapper that starts `soffice.bin`. A child started with the caller's stdout and stderr
hands those pipes to `soffice.bin`, and if `soffice.bin` outlives the script (a shutdown call that
times out kills only the wrapper), the caller's read of the script's output never sees end of file.
The .NET tests read these scripts' output, so one orphan hung a whole test run (issue #1158).

Every soffice process here therefore gets /dev/null for its standard streams and its own process
group, and is stopped by signalling the whole group, `soffice.bin` included.
"""
import os
import signal
import subprocess
import time

_DETACHED = dict(stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                 start_new_session=True)


def start_listener(port, profile, executable="soffice"):
    """Start a headless soffice that accepts UNO connections on `port`, using the profile URL `profile`."""
    return subprocess.Popen([
        executable, "--headless", "--invisible", "--norestore", "--nologo",
        "--nofirststartwizard", "--nodefault",
        "--accept=socket,host=localhost,port=%s;urp;StarOffice.ComponentContext" % port,
        "-env:UserInstallation=" + profile,
    ], **_DETACHED)


def run_detached(arguments, timeout):
    """Run a command without the caller's pipes; after `timeout` seconds, kill its whole process group."""
    try:
        process = subprocess.Popen(arguments, **_DETACHED)
    except OSError:
        return
    try:
        process.wait(timeout=timeout)
    except subprocess.TimeoutExpired:
        _stop_group(process, grace=2)


def unaccept(port, profile, executable="soffice", timeout=10):
    """Ask the listener on `port` to stop accepting connections, as the scripts' shutdown always has."""
    run_detached([executable, "--headless", "-env:UserInstallation=" + profile,
                  "--unaccept=socket,host=localhost,port=%s;urp;" % port], timeout)


def stop_listener(process, grace=10):
    """Stop a listener from `start_listener` and everything it started."""
    _stop_group(process, grace)


def _stop_group(process, grace):
    # The group outlives the wrapper whose pid names it (soffice.bin keeps running after the
    # wrapper exits), so the grace period waits for the whole group, not just the wrapper.
    _signal_group(process, signal.SIGTERM)
    deadline = time.monotonic() + grace
    while _group_alive(process) and time.monotonic() < deadline:
        time.sleep(0.05)
    _signal_group(process, signal.SIGKILL)
    try:
        process.wait(timeout=grace)
    except subprocess.TimeoutExpired:
        pass


def _group_alive(process):
    process.poll()  # reap the wrapper, so a zombie leader does not count as alive
    if not hasattr(os, "killpg"):
        return process.returncode is None
    try:
        os.killpg(process.pid, 0)
        return True
    except (ProcessLookupError, PermissionError):
        return False


def _signal_group(process, signum):
    if not hasattr(os, "killpg"):  # no process groups (Windows): signal the process alone
        if process.poll() is None:
            process.kill()
        return
    try:
        os.killpg(process.pid, signum)
    except (ProcessLookupError, PermissionError):
        pass
