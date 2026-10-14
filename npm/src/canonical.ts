/** Shared canonical-JSON serialization, used by both the export orchestrator and the browser
 * font runtime wherever a value needs a stable digest or a deterministic wire form. */

export function isWellFormedUnicode(value: string): boolean {
  for (let index = 0; index < value.length; index++) {
    const unit = value.charCodeAt(index);
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const next = value.charCodeAt(++index);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false;
    } else if (unit >= 0xdc00 && unit <= 0xdfff) {
      return false;
    }
  }
  return true;
}

export function assertWellFormedUnicode(value: string): void {
  if (!isWellFormedUnicode(value)) {
    throw new TypeError("Canonical JSON does not support unpaired UTF-16 surrogates");
  }
}

function canonicalValue(value: unknown): unknown {
  if (value === null || typeof value === "boolean") return value;
  if (typeof value === "string") {
    assertWellFormedUnicode(value);
    return value;
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw new TypeError("Canonical JSON does not support non-finite numbers");
    return Object.is(value, -0) ? 0 : value;
  }
  if (Array.isArray(value)) return value.map(canonicalValue);
  if (typeof value === "object") {
    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) {
      throw new TypeError("Canonical JSON supports only plain objects");
    }
    const record = value as Record<string, unknown>;
    // A normal object would treat `__proto__` as a prototype mutation rather than an own data
    // property, silently dropping it from the canonical payload. The exporter's copy
    // (`npm-export/src/canonical.ts`) does the same, so both sides digest the same bytes.
    const result: Record<string, unknown> = Object.create(null) as Record<string, unknown>;
    for (const key of Object.keys(record).sort()) {
      assertWellFormedUnicode(key);
      const member = record[key];
      if (member !== undefined) result[key] = canonicalValue(member);
    }
    return result;
  }
  throw new TypeError(`Canonical JSON does not support ${typeof value}`);
}

/** Serialize report/schema values with recursively sorted object keys and no insignificant space. */
export function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalValue(value));
}
