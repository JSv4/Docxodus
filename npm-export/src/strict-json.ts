/** Parse JSON while rejecting duplicate properties and pathological nesting. */
export function strictJsonParse(
  source: string,
  onFailure: (detail: string) => never = (detail) => { throw new SyntaxError(detail); },
): unknown {
  let cursor = 0;
  const wellFormed = (value: string): boolean => {
    for (let index = 0; index < value.length; index++) {
      const unit = value.charCodeAt(index);
      if (unit >= 0xd800 && unit <= 0xdbff) {
        const next = value.charCodeAt(++index);
        if (!(next >= 0xdc00 && next <= 0xdfff)) return false;
      } else if (unit >= 0xdc00 && unit <= 0xdfff) return false;
    }
    return true;
  };
  const whitespace = () => {
    while (cursor < source.length && /[\u0009\u000a\u000d\u0020]/.test(source[cursor])) cursor++;
  };
  const parseStringToken = (): string => {
    const start = cursor;
    if (source[cursor++] !== '"') throw new SyntaxError(`Expected string at ${start}`);
    while (cursor < source.length) {
      const character = source[cursor++];
      if (character === '"') {
        const value = JSON.parse(source.slice(start, cursor)) as string;
        if (!wellFormed(value)) throw new SyntaxError(`Unpaired Unicode surrogate at ${start}`);
        return value;
      }
      if (character.charCodeAt(0) < 0x20) {
        throw new SyntaxError(`Control character at ${cursor - 1}`);
      }
      if (character !== "\\") continue;
      if (cursor >= source.length) throw new SyntaxError("Unterminated JSON escape");
      const escape = source[cursor++];
      if (escape === "u") {
        if (!/^[0-9a-fA-F]{4}$/.test(source.slice(cursor, cursor + 4))) {
          throw new SyntaxError(`Invalid Unicode escape at ${cursor - 2}`);
        }
        cursor += 4;
      } else if (!'"\\/bfnrt'.includes(escape)) {
        throw new SyntaxError(`Invalid JSON escape at ${cursor - 2}`);
      }
    }
    throw new SyntaxError("Unterminated JSON string");
  };
  const parseValue = (depth: number): void => {
    if (depth > 128) throw new SyntaxError("JSON nesting is too deep");
    whitespace();
    const character = source[cursor];
    if (character === '"') {
      parseStringToken();
      return;
    }
    if (character === "{") {
      cursor++;
      whitespace();
      const keys = new Set<string>();
      if (source[cursor] === "}") { cursor++; return; }
      while (cursor < source.length) {
        whitespace();
        const key = parseStringToken();
        if (keys.has(key)) throw new SyntaxError(`Duplicate JSON property ${JSON.stringify(key)}`);
        keys.add(key);
        whitespace();
        if (source[cursor++] !== ":") throw new SyntaxError(`Expected colon at ${cursor - 1}`);
        parseValue(depth + 1);
        whitespace();
        const separator = source[cursor++];
        if (separator === "}") return;
        if (separator !== ",") throw new SyntaxError(`Expected object separator at ${cursor - 1}`);
      }
      throw new SyntaxError("Unterminated JSON object");
    }
    if (character === "[") {
      cursor++;
      whitespace();
      if (source[cursor] === "]") { cursor++; return; }
      while (cursor < source.length) {
        parseValue(depth + 1);
        whitespace();
        const separator = source[cursor++];
        if (separator === "]") return;
        if (separator !== ",") throw new SyntaxError(`Expected array separator at ${cursor - 1}`);
      }
      throw new SyntaxError("Unterminated JSON array");
    }
    const rest = source.slice(cursor);
    const keyword = /^(?:true|false|null)/.exec(rest)?.[0];
    const number = keyword === undefined
      ? /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?/.exec(rest)?.[0]
      : undefined;
    if (number !== undefined && !Number.isFinite(Number(number))) {
      throw new SyntaxError(`Non-finite JSON number at ${cursor}`);
    }
    const literal = keyword ?? number;
    if (!literal) throw new SyntaxError(`Invalid JSON value at ${cursor}`);
    cursor += literal.length;
  };
  try {
    parseValue(0);
    whitespace();
    if (cursor !== source.length) throw new SyntaxError(`Trailing JSON data at ${cursor}`);
    return JSON.parse(source) as unknown;
  } catch (error) {
    return onFailure(error instanceof Error ? error.message : String(error));
  }
}

export function decodeStrictUtf8(bytes: Uint8Array, label: string): string {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch (cause) {
    throw new TypeError(`${label} is not strict UTF-8.`, { cause });
  }
}
