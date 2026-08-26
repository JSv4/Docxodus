import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { dirname, join } from "node:path";
import { describe, test } from "node:test";
import { fileURLToPath } from "node:url";
import {
  checkExportEnvironment,
  evaluateExportEnvironment,
} from "../dist/environment.js";

const here = dirname(fileURLToPath(import.meta.url));
const packageRoot = dirname(here);
const cliEntry = join(packageRoot, "dist", "cli.js");

// The judgment is a pure function over probed facts, so the cases live hosts cannot
// stage (root, denied namespaces) are pinned here directly (issue #595).
describe("evaluateExportEnvironment", () => {
  const clean = Object.freeze({
    platform: "linux",
    effectiveUserId: 1000,
    userNamespaces: "available",
    browserExecutable: "resolved",
  });

  test("a clean unprivileged Linux host is ok with zero findings", () => {
    const report = evaluateExportEnvironment({ ...clean });
    assert.equal(report.ok, true);
    assert.deepEqual(report.findings, []);
  });

  test("root is fatal even when user namespaces are available", () => {
    const report = evaluateExportEnvironment({ ...clean, effectiveUserId: 0 });
    assert.equal(report.ok, false);
    const finding = report.findings.find((entry) => entry.code === "running_as_root");
    assert.equal(finding?.severity, "fatal");
    assert.match(finding.remediation, /unprivileged user/);
    // The root guidance must not send the operator to the namespace knob.
    assert.doesNotMatch(finding.remediation, /apparmor|userns|namespace/i);
  });

  test("denied user namespaces are fatal on Linux with the AppArmor knob named", () => {
    const report = evaluateExportEnvironment({ ...clean, userNamespaces: "unavailable" });
    assert.equal(report.ok, false);
    const finding = report.findings.find((entry) => entry.code === "user_namespaces_unavailable");
    assert.equal(finding?.severity, "fatal");
    assert.match(finding.remediation, /apparmor_restrict_unprivileged_userns/);
    assert.match(finding.remediation, /CLONE_NEWUSER/);
  });

  test("an unrunnable namespace probe is advisory, not fatal", () => {
    const report = evaluateExportEnvironment({ ...clean, userNamespaces: "unknown" });
    assert.equal(report.ok, true);
    assert.equal(report.findings[0]?.code, "user_namespaces_unknown");
    assert.equal(report.findings[0]?.severity, "advisory");
  });

  test("namespace findings are Linux-only; the root check is not", () => {
    const mac = evaluateExportEnvironment({
      platform: "darwin",
      effectiveUserId: 0,
      userNamespaces: "unknown",
      browserExecutable: "resolved",
    });
    assert.deepEqual(mac.findings.map((entry) => entry.code), ["running_as_root"]);
  });

  test("an unresolvable browser executable is fatal with install guidance", () => {
    const report = evaluateExportEnvironment({ ...clean, browserExecutable: "missing" });
    assert.equal(report.ok, false);
    const finding = report.findings.find((entry) => entry.code === "browser_executable_missing");
    assert.match(finding?.remediation ?? "", /@playwright\/browser-chromium|browserExecutablePath/);
  });
});

describe("checkExportEnvironment (live probes)", () => {
  test("produces a well-formed report whose ok matches its findings", async () => {
    const report = await checkExportEnvironment();
    assert.equal(typeof report.ok, "boolean");
    for (const finding of report.findings) {
      assert.equal(typeof finding.code, "string");
      assert.ok(finding.message.length > 0);
      assert.ok(finding.remediation.length > 0);
    }
    assert.equal(report.ok, report.findings.every((entry) => entry.severity !== "fatal"));
  });

  test("a bogus explicit browser path is reported missing", async () => {
    const report = await checkExportEnvironment({
      browserExecutablePath: join(packageRoot, "does-not-exist", "chromium"),
    });
    assert.equal(report.ok, false);
    assert.ok(report.findings.some((entry) => entry.code === "browser_executable_missing"));
  });
});

describe("docxodus doctor", () => {
  test("reports the missing browser with exit 1 and remediation on stderr", () => {
    const result = spawnSync(process.execPath, [
      cliEntry,
      "doctor",
      "--browser-executable", join(packageRoot, "does-not-exist", "chromium"),
    ], { cwd: packageRoot });
    assert.equal(result.status, 1);
    const stderr = result.stderr.toString();
    assert.match(stderr, /browser_executable_missing/);
    assert.match(stderr, /not ok/);
  });
});
