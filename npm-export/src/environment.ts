import { spawn } from "node:child_process";
import { stat } from "node:fs/promises";
import { chromium } from "playwright-core";

/**
 * Deployment preflight (issue #595). The export refuses to launch Chromium without its OS
 * sandbox, which turns two host properties most container defaults violate into launch-time
 * failures: the process must not run as root (Chromium's sandbox refuses root even where user
 * namespaces are fully enabled), and unprivileged user namespaces must be permitted. Probing
 * both at boot — plus whether a Chromium executable is resolvable at all — lets a deployment
 * fail with configuration guidance instead of at its first user-facing conversion.
 */

export type ExportEnvironmentFindingCode =
  | "running_as_root"
  | "user_namespaces_unavailable"
  | "user_namespaces_unknown"
  | "browser_executable_missing";

export interface ExportEnvironmentFinding {
  code: ExportEnvironmentFindingCode;
  /** "fatal" findings make {@link ExportEnvironmentReport.ok} false; "advisory" ones do not. */
  severity: "fatal" | "advisory";
  message: string;
  remediation: string;
}

export interface ExportEnvironmentReport {
  /** True when no fatal finding was raised — exports are expected to launch on this host. */
  ok: boolean;
  findings: ExportEnvironmentFinding[];
}

/**
 * The observed host facts {@link evaluateExportEnvironment} judges. Split from the probing so
 * the judgment is testable on hosts where the conditions cannot be created (CI runs unprivileged
 * with namespaces enabled; nothing there can observe the root or denied-namespace cases live).
 */
export interface ExportEnvironmentProbes {
  platform: NodeJS.Platform;
  /** `process.geteuid()` where the platform has it; undefined on Windows. */
  effectiveUserId: number | undefined;
  /** Result of the `unshare --user --map-root-user true` probe; "unknown" when it could not run. */
  userNamespaces: "available" | "unavailable" | "unknown";
  browserExecutable: "resolved" | "missing";
}

/** Pure judgment over probed host facts — the semantics behind {@link checkExportEnvironment}. */
export function evaluateExportEnvironment(probes: ExportEnvironmentProbes): ExportEnvironmentReport {
  const findings: ExportEnvironmentFinding[] = [];

  if (probes.effectiveUserId === 0) {
    findings.push({
      code: "running_as_root",
      severity: "fatal",
      message: "This process runs as root; Chromium's sandbox cannot be used by root, so the "
        + "export's browser launch will be refused even where unprivileged user namespaces are "
        + "fully enabled.",
      remediation: "Run the export as an unprivileged user — in Docker set a non-root USER (or "
        + "runAsNonRoot/runAsUser in a Kubernetes securityContext). The export runtime never "
        + "launches Chromium without its process sandbox.",
    });
  }

  if (probes.platform === "linux" && probes.userNamespaces === "unavailable") {
    findings.push({
      code: "user_namespaces_unavailable",
      severity: "fatal",
      message: "This host denies unprivileged user namespaces, which Chromium's process sandbox "
        + "requires.",
      remediation: "Permit unprivileged user namespaces on the render host — for example "
        + "kernel.apparmor_restrict_unprivileged_userns=0 on Ubuntu 23.10 and later; in "
        + "containers, use a seccomp profile that permits clone with CLONE_NEWUSER and do not "
        + "drop the capabilities the namespace path needs.",
    });
  }

  if (probes.platform === "linux" && probes.userNamespaces === "unknown") {
    findings.push({
      code: "user_namespaces_unknown",
      severity: "advisory",
      message: "Whether this host permits unprivileged user namespaces could not be determined "
        + "(the unshare probe did not run).",
      remediation: "Verify manually: `unshare --user --map-root-user true` must exit 0 for "
        + "Chromium's sandbox to work.",
    });
  }

  if (probes.browserExecutable === "missing") {
    findings.push({
      code: "browser_executable_missing",
      severity: "fatal",
      message: "No Chromium executable is resolvable for the export.",
      remediation: "Install @playwright/browser-chromium during deployment or provide "
        + "browserExecutablePath (CLI: --browser-executable or DOCXODUS_CHROMIUM_PATH).",
    });
  }

  return { ok: findings.every((finding) => finding.severity !== "fatal"), findings };
}

async function probeUserNamespaces(timeoutMs: number): Promise<"available" | "unavailable" | "unknown"> {
  return new Promise((resolvePromise) => {
    let settled = false;
    const settle = (value: "available" | "unavailable" | "unknown") => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolvePromise(value);
    };
    let child: ReturnType<typeof spawn>;
    try {
      child = spawn("unshare", ["--user", "--map-root-user", "true"], {
        stdio: "ignore",
        timeout: timeoutMs,
      });
    } catch {
      resolvePromise("unknown");
      return;
    }
    const timer = setTimeout(() => {
      child.kill();
      settle("unknown");
    }, timeoutMs);
    child.on("error", () => settle("unknown"));
    child.on("exit", (code, sig) => settle(sig !== null ? "unknown" : code === 0 ? "available" : "unavailable"));
  });
}

async function probeBrowserExecutable(explicitPath: string | undefined): Promise<"resolved" | "missing"> {
  try {
    const path = explicitPath ?? chromium.executablePath();
    if (!path) return "missing";
    const identity = await stat(path);
    return identity.isFile() ? "resolved" : "missing";
  } catch {
    return "missing";
  }
}

/**
 * Probe this host and judge whether exports are expected to launch (issue #595): the process
 * must not be root, unprivileged user namespaces must be permitted (Linux), and a Chromium
 * executable must resolve. Run it at deployment boot — or via `docxodus doctor` — so
 * misconfiguration surfaces as guidance instead of a first-conversion failure. Read-only:
 * nothing is launched, installed, or modified beyond one `unshare … true` no-op probe.
 */
export async function checkExportEnvironment(
  options?: { browserExecutablePath?: string },
): Promise<ExportEnvironmentReport> {
  const platform = process.platform;
  return evaluateExportEnvironment({
    platform,
    effectiveUserId: typeof process.geteuid === "function" ? process.geteuid() : undefined,
    userNamespaces: platform === "linux" ? await probeUserNamespaces(5_000) : "unknown",
    browserExecutable: await probeBrowserExecutable(options?.browserExecutablePath),
  });
}
