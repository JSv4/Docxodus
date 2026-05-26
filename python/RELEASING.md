# Releasing `docx-scalpel` to PyPI

`docx-scalpel` is published via **PyPI Trusted Publishing (OIDC)** — no API tokens stored in GitHub secrets. The publish job in [`.github/workflows/python-publish.yml`](../.github/workflows/python-publish.yml) exchanges an OIDC token signed by GitHub for a short-lived PyPI upload credential at publish time.

## One-time setup (repo-level)

Two things must exist before the first publish:

### 1. GitHub Actions environments

Create two environments in **Settings → Environments**:

| Name | Purpose | Suggested protection |
|---|---|---|
| `testpypi` | Dry-run publishes via `workflow_dispatch` | None — used freely |
| `pypi` | Production publishes (auto on `docx-scalpel-v*` tag push) | Required reviewer, restrict to `main` branch |

The names are case-sensitive and must match the `environment.name` fields in the workflow.

### 2. PyPI / TestPyPI trusted publishers

For **each** of TestPyPI and PyPI, register a pending trusted publisher (the `docx-scalpel` project doesn't exist yet on first upload).

- TestPyPI: https://test.pypi.org/manage/account/publishing/
- PyPI: https://pypi.org/manage/account/publishing/

Fill out the form exactly:

| Field | Value |
|---|---|
| PyPI project name | `docx-scalpel` |
| Owner | `JSv4` |
| Repository name | `Docxodus` |
| Workflow filename | `python-publish.yml` |
| Environment name | `testpypi` (or `pypi`, matching the row) |

After the first successful publish, the "pending" publisher becomes a regular trusted publisher on the project page.

## Routine release

**docx-scalpel ships on its own tag pattern, decoupled from Docxodus core / npm / binaries.**

| Tag pattern | Workflow | Publishes |
|---|---|---|
| `v*` (e.g. `v1.2.3`) | `publish.yml` | NuGet (Docxodus + redline + docx2html + docx2oc) + npm (`docxodus`) + GitHub Release binaries |
| `docx-scalpel-v*` (e.g. `docx-scalpel-v0.1.0a1`) | `python-publish.yml` | PyPI (`docx-scalpel`) |

The two tag namespaces are independent. A docx-scalpel point-release doesn't drag Docxodus core through a version bump, and a Docxodus core release doesn't force unrelated PyPI churn. Each docx-scalpel wheel bundles a `docxodus-pyhost` built from the same commit the tag points to, so there's no runtime version-pinning between the two distributions.

### Recommended sequence

1. **Dry-run on TestPyPI** (manual)
   - Actions → `python-publish` → Run workflow
   - Branch: whatever you're cutting from (default branch for stable, feature branch for testing the pipeline itself)
   - `version`: PEP 440 string, e.g. `0.1.0a1`
   - `target`: `testpypi`
   - Confirm the upload by installing in a fresh venv:
     ```bash
     pip install --index-url https://test.pypi.org/simple/ \
         --extra-index-url https://pypi.org/simple/ \
         docx-scalpel==0.1.0a1
     python -c "from docx_scalpel import ping; print(ping())"
     ```

2. **Real release on PyPI** (tag-driven)
   - Bump the version in `python/pyproject.toml`. CI overwrites this at build time, but keeping the in-tree value current is helpful for editable installs and PR readability.
     ```bash
     # python/pyproject.toml: version = "0.1.0a1"
     git commit -am "chore(python): bump docx-scalpel to 0.1.0a1"
     git tag docx-scalpel-v0.1.0a1
     git push origin main docx-scalpel-v0.1.0a1
     ```
   - The tag fires `python-publish.yml`. `resolve` extracts `0.1.0a1` from `${GITHUB_REF#refs/tags/docx-scalpel-v}` and resolves target = `pypi`.

### Pre-releases

PyPI accepts PEP 440 pre-release identifiers — `0.1.0a1`, `0.1.0b2`, `0.1.0rc1`, `0.1.0.dev3`. Use the tag form `docx-scalpel-v0.1.0a1` and pip will only install pre-releases when `--pre` is passed.

## Wheel scope

The workflow ships one wheel per RID plus a single sdist:

| RID | Runner | Wheel platform tag | Binary |
|---|---|---|---|
| `linux-x64` | `ubuntu-latest` | `manylinux_2_28_x86_64` | `docxodus-pyhost` |
| `linux-arm64` | `ubuntu-22.04-arm` | `manylinux_2_28_aarch64` | `docxodus-pyhost` |
| `osx-x64` | `macos-13` | `macosx_11_0_x86_64` | `docxodus-pyhost` |
| `osx-arm64` | `macos-14` | `macosx_11_0_arm64` | `docxodus-pyhost` |
| `win-x64` | `windows-latest` | `win_amd64` | `docxodus-pyhost.exe` |

The matrix is in `.github/workflows/python-publish.yml` under `build-wheel`. Each entry runs the same step sequence (publish .NET host → stage → smoke-test → build wheel → retag → lifecycle-test from a fresh-venv install). `fail-fast: false` so all RID failures surface at once, but `publish` still gates on every matrix entry succeeding.

`docx_scalpel-${VER}.tar.gz` (the sdist) has **no bundled binary** — install requires `DOCXODUS_HOST` set or a Docxodus monorepo clone with the host built locally.

### Known gaps

- **Code signing on macOS** — wheels ship an unsigned `docxodus-pyhost`. First launch may trigger a Gatekeeper warning; user can bypass via right-click → Open or `xattr -d com.apple.quarantine`. Proper signing + notarization is a future ask; needs an Apple Developer ID and an `APPLE_*` secret bundle in CI.
- **Authenticode signing on Windows** — same story. Unsigned `.exe` triggers SmartScreen on first run. Needs a code-signing cert.
- **Glibc compliance for Linux ARM** — `ubuntu-22.04-arm` runner has glibc 2.35; we claim `manylinux_2_28_aarch64`. The .NET 8 self-contained binary should be glibc-2.28-compatible but we don't run `auditwheel` to enforce it. If users on older arm64 distros report `GLIBC_2.34 not found`, build inside a `quay.io/pypa/manylinux_2_28_aarch64` container instead.

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| `OIDC token exchange failed` in publish step | Trusted publisher not registered, or the environment name in the workflow doesn't match what's configured on PyPI |
| `403 Forbidden` from PyPI | Same as above, or the project name was claimed by someone else after the pending-publisher form was filled out |
| `pyhost ping fails in CI` smoke test | `dotnet publish` succeeded but produced a binary missing a runtime dep; check the .NET SDK version in `setup-dotnet` matches the one in `global.json` |
| Wheel installs but `find_host()` raises `DocxodusHostNotFoundError` | Wheel was built without the binary staged in `src/docx_scalpel/_bin/`; the workflow's "Stage binary into wheel layout" step did not run or wrote to the wrong path |
| `auditwheel` complains about glibc | Build on a manylinux container (`quay.io/pypa/manylinux_2_28_x86_64`) instead of the `ubuntu-latest` runner |

## Why trusted publishing

- No long-lived API tokens to rotate or leak via a compromised GitHub Secret.
- Token is scoped to one workflow run on one repo on one environment — exfiltration window is minutes, not the lifetime of a secret.
- No password to share among maintainers; access is governed by who can push tags / approve environment deployments.

See [PyPI's Trusted Publishers docs](https://docs.pypi.org/trusted-publishers/) for the full rationale and security model.
