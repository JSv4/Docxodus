function assertWellFormedUnicode(value: string): void {
  for (let index = 0; index < value.length; index++) {
    const unit = value.charCodeAt(index);
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const next = value.charCodeAt(++index);
      if (!(next >= 0xdc00 && next <= 0xdfff)) {
        throw new TypeError('Canonical JSON rejects unpaired UTF-16 surrogates');
      }
    } else if (unit >= 0xdc00 && unit <= 0xdfff) {
      throw new TypeError('Canonical JSON rejects unpaired UTF-16 surrogates');
    }
  }
}

function canonicalValue(value: unknown): unknown {
  if (value === null || typeof value === 'boolean') return value;
  if (typeof value === 'string') {
    assertWellFormedUnicode(value);
    return value;
  }
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new TypeError('Canonical JSON rejects non-finite numbers');
    return Object.is(value, -0) ? 0 : value;
  }
  if (Array.isArray(value)) return value.map(canonicalValue);
  if (typeof value === 'object') {
    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) {
      throw new TypeError('Canonical JSON supports only plain objects');
    }
    const source = value as Record<string, unknown>;
    const result: Record<string, unknown> = Object.create(null) as Record<string, unknown>;
    for (const key of Object.keys(source).sort()) {
      assertWellFormedUnicode(key);
      if (source[key] !== undefined) result[key] = canonicalValue(source[key]);
    }
    return result;
  }
  throw new TypeError(`Canonical JSON rejects ${typeof value} values`);
}

export function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalValue(value));
}
