"""The renderer builds rendering-fidelity.yml starts in the background record their exit status whatever it
is, so the wait step ends with that status and prints the build log instead of waiting out its timeout.

Each start step and its wait step run as GitHub Actions runs a `run:` step (`bash -eo pipefail`), with
stub `dotnet` and `python3` executables standing in for the workload install and the bench build.
"""
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest

WORKFLOW = Path(__file__).resolve().parents[2] / '.github' / 'workflows' / 'rendering-fidelity.yml'
BASH = shutil.which('bash')
STUB = '#!/bin/sh\necho "{name} $*"\nexit "${{{status}:-0}}"\n'


def run_block(step):
    """The `run: |` script of the workflow step named `step`."""
    lines = WORKFLOW.read_text(encoding='utf-8').splitlines()
    at = next(i for i, line in enumerate(lines) if line.strip() == f'- name: {step}')
    run = next(i for i in range(at + 1, len(lines)) if lines[i].strip() == 'run: |')
    indent = len(lines[run]) - len(lines[run].lstrip())
    block = []
    for line in lines[run + 1:]:
        if line.strip() and len(line) - len(line.lstrip()) <= indent:
            break
        block.append(line)
    return '\n'.join(block) + '\n'


@unittest.skipUnless(BASH, 'needs bash')
class BackgroundRendererBuild(unittest.TestCase):
    def wait_after_build(self, build, statuses):
        """Runs `build`'s start step, then its wait step; returns the wait step's exit code and output."""
        with tempfile.TemporaryDirectory() as root:
            root = Path(root)
            bin_dir, work, temp = root / 'bin', root / 'work', root / 'temp'
            for d in (bin_dir, work / 'oracle', temp):
                d.mkdir(parents=True)
            for name, status in (('dotnet', 'DOTNET_STATUS'), ('python3', 'BENCH_STATUS')):
                (bin_dir / name).write_text(STUB.format(name=name, status=status), newline='\n')
                (bin_dir / name).chmod(0o755)
            env = dict(os.environ, PATH=f'{bin_dir}{os.pathsep}{os.environ["PATH"]}', RUNNER_TEMP=temp.as_posix(),
                       MAIN_COMMIT='main', CANDIDATE_COMMIT='candidate', **statuses)

            def step(name, timeout):
                script = root / 'step.sh'
                script.write_text(run_block(name), newline='\n')
                return subprocess.run([BASH, '--noprofile', '--norc', '-eo', 'pipefail', script.as_posix()],
                                      cwd=work, env=env, capture_output=True, text=True, timeout=timeout)

            self.assertEqual(step(f'Start the {build} renderer build', 30).returncode, 0)
            try:
                waited = step(f'Wait for the {build} renderer build', 20)
            except subprocess.TimeoutExpired:
                self.fail(f'the {build} build never wrote build.status, so the wait step ran into its timeout')
            return waited.returncode, waited.stdout

    def test_a_failed_build_ends_the_wait_with_its_status_and_log(self):
        for build in ('main', 'candidate'):
            with self.subTest(build=build, failing='bench build'):
                code, log = self.wait_after_build(build, {'BENCH_STATUS': '2'})
                self.assertEqual(code, 2)
                self.assertIn('python3 -m bench build-docxodus', log)
            with self.subTest(build=build, failing='workload install'):
                code, log = self.wait_after_build(build, {'DOTNET_STATUS': '3'})
                self.assertEqual(code, 3)
                self.assertIn('dotnet workload install wasm-tools', log)
                self.assertNotIn('bench build-docxodus', log, 'the build stops at its first failing command')

    def test_a_successful_build_still_continues(self):
        for build in ('main', 'candidate'):
            with self.subTest(build=build):
                code, log = self.wait_after_build(build, {})
                self.assertEqual(code, 0)
                self.assertIn(f'python3 -m bench build-docxodus --source ../{build}-source --commit {build}', log)


if __name__ == '__main__':
    unittest.main()
