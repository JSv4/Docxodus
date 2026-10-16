import os
from pathlib import Path
import signal
import subprocess
import sys
import tempfile
import textwrap
import time
import unittest

import lo_soffice

HERE = Path(__file__).resolve().parent


class SofficeProcessTests(unittest.TestCase):
    """A soffice process that outlives its script must not hold the script's output open (issue #1158)."""

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.pid_file = self.folder / "grandchild.pid"
        # Stands in for the soffice wrapper: it starts a long-lived child (soffice.bin's role), which
        # inherits whatever standard streams the wrapper was given, then exits.
        self.fake_soffice = self.folder / "fake_soffice.py"
        self.fake_soffice.write_text(f"#!{sys.executable}\n" + textwrap.dedent(f"""\
            import subprocess, sys
            child = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(60)"])
            open({str(self.pid_file)!r}, "w").write(str(child.pid))
            """))
        self.fake_soffice.chmod(0o755)
        self.addCleanup(self.kill_grandchild)

    def kill_grandchild(self):
        try:
            os.kill(int(self.pid_file.read_text()), signal.SIGKILL)
        except (FileNotFoundError, ValueError, ProcessLookupError):
            pass

    def wait_for_pid_file(self):
        deadline = time.monotonic() + 10
        while not self.pid_file.exists() or not self.pid_file.read_text():
            self.assertLess(time.monotonic(), deadline, "the fake soffice never started its child")
            time.sleep(0.05)
        return int(self.pid_file.read_text())

    def test_listener_does_not_inherit_the_callers_output_pipe(self):
        script = textwrap.dedent(f"""\
            import sys
            sys.path.insert(0, {str(HERE)!r})
            import lo_soffice
            lo_soffice.start_listener("0", "file:///nowhere", executable={str(self.fake_soffice)!r})
            print("RESULT: OK", flush=True)
            """)
        caller = subprocess.Popen([sys.executable, "-c", script], stdout=subprocess.PIPE,
                                  stderr=subprocess.PIPE, text=True)
        # Reading to end of file is what the .NET tests do. It returns only if no descendant kept the pipe.
        stdout, _ = caller.communicate(timeout=10)
        self.wait_for_pid_file()
        self.assertEqual("RESULT: OK\n", stdout)

    def test_stop_listener_ends_what_the_listener_started(self):
        listener = lo_soffice.start_listener("0", "file:///nowhere",
                                             executable=str(self.fake_soffice))
        grandchild = self.wait_for_pid_file()
        lo_soffice.stop_listener(listener, grace=2)
        self.assert_gone(grandchild)

    def test_run_detached_kills_the_process_group_at_the_timeout(self):
        started = time.monotonic()
        lo_soffice.run_detached([sys.executable, "-c", textwrap.dedent(f"""\
            import subprocess, sys, time
            child = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(60)"])
            open({str(self.pid_file)!r}, "w").write(str(child.pid))
            time.sleep(60)
            """)], timeout=1)
        self.assertLess(time.monotonic() - started, 10)
        self.assert_gone(self.wait_for_pid_file())

    def assert_gone(self, pid):
        deadline = time.monotonic() + 5
        while time.monotonic() < deadline:
            try:
                state = (Path("/proc") / str(pid) / "stat").read_text().split()[2]
            except (FileNotFoundError, ProcessLookupError):
                return
            if state in ("Z", "X"):
                return
            time.sleep(0.05)
        self.fail("process %d is still running" % pid)


if __name__ == "__main__":
    unittest.main()
