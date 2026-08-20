import { DocxodusExportError } from "./contracts.js";

/**
 * Operator diagnostics for stderr.
 *
 * Writing a cause to a terminal is not serializing it: nothing here reaches `detail`, `toJSON()`,
 * or the render report, which stay free of the cause and stack material a Node error may carry.
 */

const MAX_HUMAN_DIAGNOSTIC_CHARACTERS = 16_384;
const MAX_HUMAN_CAUSE_CHARACTERS = 8_192;
const MAX_CAUSE_DEPTH = 8;

/** Escapes are matched whole, so no bracket or parameter residue survives the strip. */
const ANSI_OPERATING_SYSTEM_COMMAND = /\u001b\][\s\S]*?(?:\u0007|\u001b\\)/g;
const ANSI_CONTROL_SEQUENCE = /\u001b\[[\u0020-\u003f]*[\u0040-\u007e]/g;
/** Every C0/C1 code point and DEL except the newline this rendering uses as its separator. */
const FORBIDDEN_CONTROL = /[\u0000-\u0009\u000b-\u001f\u007f-\u009f]/g;

/**
 * Renders text a terminal can print. Newline survives deliberately: the diagnostics this exists to
 * surface — Chromium's launch log above all — are multi-line, and flattening them would restore
 * exactly the unreadability being fixed.
 */
function terminalSafeText(text: string): string {
  return text
    .replace(ANSI_OPERATING_SYSTEM_COMMAND, "")
    .replace(ANSI_CONTROL_SEQUENCE, "")
    .replace(/\r\n?/g, "\n")
    .replace(/\t/g, " ")
    .replace(FORBIDDEN_CONTROL, "");
}

function boundedText(text: string, maximum: number): string {
  if (maximum <= 0) return "";
  if (text.length <= maximum) return text;
  return maximum <= 3 ? text.slice(0, maximum) : `${text.slice(0, maximum - 3)}...`;
}

function errorText(value: unknown): string {
  try {
    return value instanceof Error ? value.message : String(value);
  } catch {
    return "an unrenderable diagnostic value";
  }
}

function walk(root: unknown, includeRoot: boolean, maxDepth: number): string[] {
  const seen = new Set<unknown>();
  const messages: string[] = [];
  const visit = (value: unknown, depth: number): void => {
    if (value === undefined || value === null || depth > maxDepth) return;
    if (typeof value === "object" || typeof value === "function") {
      if (seen.has(value)) return;
      seen.add(value);
    }
    if (depth > 0 || includeRoot) {
      const text = errorText(value);
      // `new AggregateError([...])` carries no message of its own; only its members say anything.
      if (text.length > 0) messages.push(text);
    }
    if (value instanceof AggregateError && Array.isArray(value.errors)) {
      for (const entry of value.errors) visit(entry, depth + 1);
    }
    if (value instanceof Error && value.cause !== undefined) visit(value.cause, depth + 1);
  };
  visit(root, 0);
  return messages;
}

/** Every message in the chain, the root's own included. */
export function errorMessages(value: unknown, maxDepth: number = MAX_CAUSE_DEPTH): string[] {
  return walk(value, true, maxDepth);
}

/**
 * Everything beneath `error`: `cause` chains and `AggregateError` members alike, because the browser
 * launch path wraps a primary failure together with a cleanup failure. The root's own message is
 * omitted, since callers render it themselves.
 */
function causeMessages(error: unknown, maxDepth: number = MAX_CAUSE_DEPTH): string[] {
  return walk(error, false, maxDepth);
}

export function humanDiagnostic(error: unknown): string {
  let message: string;
  if (error instanceof DocxodusExportError) {
    message = `${error.code} (${error.phase}): ${error.message}\nRemediation: ${error.remediation}`
      + (error.detail ? `\nDetail: ${error.detail}` : "")
      + (error.committedDestinations.length
        ? `\nAlready committed: ${error.committedDestinations.join(", ")}`
        : "");
  } else {
    message = errorText(error);
  }
  const causes = causeMessages(error).map((text) => `Cause: ${terminalSafeText(text)}`);
  // The cause gets its own budget before the rest is bounded, so a long detail cannot truncate away
  // the one line that explains the failure.
  const causeBlock = causes.length === 0
    ? ""
    : `\n${boundedText(causes.join("\n"), MAX_HUMAN_CAUSE_CHARACTERS)}`;
  return boundedText(
    terminalSafeText(message),
    MAX_HUMAN_DIAGNOSTIC_CHARACTERS - causeBlock.length,
  ) + causeBlock;
}
