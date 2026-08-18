/** Canonical standalone automatic-resource policy shared by sanitization and readiness. */

const STANDALONE_DATA_MEDIA_TYPES = new Set([
  "image/png", "image/jpeg", "image/gif", "image/bmp", "image/webp", "image/tiff",
  "image/x-icon", "font/woff", "font/woff2", "font/ttf", "font/otf",
  "application/font-woff", "application/vnd.ms-fontobject",
]);

export interface StandaloneDataUrlInfo {
  mediaType: string;
  byteLength: number;
}

export function dataUrlInfo(value: string): StandaloneDataUrlInfo | undefined {
  if (!value.startsWith("data:")) return undefined;
  const comma = value.indexOf(",");
  if (comma < 0) return undefined;
  const metadata = value.slice(5, comma);
  const segments = metadata.split(";");
  const mediaType = (segments.shift() ?? "").toLowerCase();
  if (mediaType === "" && value === "data:,") return { mediaType: "", byteLength: 0 };
  if (!STANDALONE_DATA_MEDIA_TYPES.has(mediaType)) return undefined;
  if (segments.some((segment) => segment.toLowerCase() !== "base64")) return undefined;
  if (!segments.some((segment) => segment.toLowerCase() === "base64")) return undefined;
  const payload = value.slice(comma + 1);
  if (!/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(payload)) {
    return undefined;
  }
  const padding = payload.endsWith("==") ? 2 : payload.endsWith("=") ? 1 : 0;
  return { mediaType, byteLength: payload.length / 4 * 3 - padding };
}

export function automaticUrlAllowed(value: string, allowFragment = false): boolean {
  const trimmed = value.trim();
  return trimmed === "" || dataUrlInfo(trimmed) !== undefined
    || (allowFragment && trimmed.startsWith("#"));
}

export function standaloneSrcsetAllowed(value: string): boolean {
  // Fail closed to one self-contained candidate. A loose comma split would
  // misparse the comma that is part of every data URL and could retain a later
  // network candidate.
  const match = /^\s*(data:\S+?)(?:\s+(?:\d+(?:\.\d+)?x|\d+w))?\s*$/i.exec(value);
  return !!match && dataUrlInfo(match[1]) !== undefined;
}

export interface CssSecurityToken {
  kind: "import" | "substitution" | "url";
  start: number;
  end: number;
  value: string;
}

function consumeCssEscape(source: string, start: number): { value: string; end: number } {
  let cursor = start + 1;
  if (cursor >= source.length) return { value: "\ufffd", end: cursor };
  if (source[cursor] === "\r" && source[cursor + 1] === "\n") {
    return { value: "", end: cursor + 2 };
  }
  if (source[cursor] === "\n" || source[cursor] === "\r" || source[cursor] === "\f") {
    return { value: "", end: cursor + 1 };
  }
  const hexStart = cursor;
  while (cursor < source.length && cursor - hexStart < 6 && /[0-9a-f]/i.test(source[cursor])) {
    cursor++;
  }
  if (cursor > hexStart) {
    const point = Number.parseInt(source.slice(hexStart, cursor), 16);
    if (/\s/.test(source[cursor] ?? "")) {
      if (source[cursor] === "\r" && source[cursor + 1] === "\n") cursor += 2;
      else cursor++;
    }
    return {
      value: point === 0 || point > 0x10ffff || (point >= 0xd800 && point <= 0xdfff)
        ? "\ufffd"
        : String.fromCodePoint(point),
      end: cursor,
    };
  }
  return { value: source[cursor], end: cursor + 1 };
}

function consumeCssName(source: string, start: number): { value: string; end: number } {
  let value = "";
  let cursor = start;
  while (cursor < source.length) {
    const character = source[cursor];
    if (character === "\\") {
      const escape = consumeCssEscape(source, cursor);
      value += escape.value;
      cursor = escape.end;
    } else if (/[a-z0-9_-]/i.test(character) || character.charCodeAt(0) >= 0x80) {
      value += character;
      cursor++;
    } else {
      break;
    }
  }
  return { value, end: cursor };
}

function consumeCssComment(source: string, start: number): number {
  const end = source.indexOf("*/", start + 2);
  return end < 0 ? source.length : end + 2;
}

function consumeCssString(source: string, start: number): number {
  const quote = source[start];
  let cursor = start + 1;
  while (cursor < source.length) {
    if (source[cursor] === quote) return cursor + 1;
    if (source[cursor] === "\n" || source[cursor] === "\r" || source[cursor] === "\f") {
      return cursor;
    }
    if (source[cursor] === "\\") cursor = consumeCssEscape(source, cursor).end;
    else cursor++;
  }
  return cursor;
}

function consumeCssFunction(source: string, openParenthesis: number): number {
  let depth = 0;
  let cursor = openParenthesis;
  while (cursor < source.length) {
    if (source.startsWith("/*", cursor)) {
      cursor = consumeCssComment(source, cursor);
    } else if (source[cursor] === "\"" || source[cursor] === "'") {
      cursor = consumeCssString(source, cursor);
    } else if (source[cursor] === "\\") {
      cursor = consumeCssEscape(source, cursor).end;
    } else if (source[cursor] === "(") {
      depth++;
      cursor++;
    } else if (source[cursor] === ")") {
      depth--;
      cursor++;
      if (depth === 0) return cursor;
    } else {
      cursor++;
    }
  }
  return cursor;
}

function decodeCssEscapedText(source: string, stripComments: boolean): string {
  let decoded = "";
  for (let cursor = 0; cursor < source.length;) {
    if (stripComments && source.startsWith("/*", cursor)) {
      cursor = consumeCssComment(source, cursor);
    } else if (source[cursor] === "\\") {
      const escape = consumeCssEscape(source, cursor);
      decoded += escape.value;
      cursor = escape.end;
    } else {
      decoded += source[cursor++];
    }
  }
  return decoded;
}

function decodeCssUrlComponent(source: string): string {
  const trimmed = source.trim();
  if (trimmed.length >= 2 && (trimmed[0] === "\"" || trimmed[0] === "'")
    && trimmed.at(-1) === trimmed[0]) {
    return decodeCssEscapedText(trimmed.slice(1, -1), false);
  }
  return decodeCssEscapedText(trimmed, true).trim();
}

/** Tokenize the security-relevant CSS grammar while honoring escapes, strings, and comments. */
export function cssSecurityTokens(css: string): CssSecurityToken[] {
  const tokens: CssSecurityToken[] = [];
  const functionStack: string[] = [];
  for (let cursor = 0; cursor < css.length;) {
    if (css.startsWith("/*", cursor)) {
      cursor = consumeCssComment(css, cursor);
      continue;
    }
    if (css[cursor] === "\"" || css[cursor] === "'") {
      const end = consumeCssString(css, cursor);
      const context = functionStack.at(-1);
      const isImageSource = context === "image-set" || context === "-webkit-image-set";
      if (isImageSource && end > cursor && css[end - 1] === css[cursor]) {
        tokens.push({
          kind: "url",
          start: cursor,
          end,
          value: decodeCssEscapedText(css.slice(cursor + 1, end - 1), false),
        });
      }
      cursor = Math.max(end, cursor + 1);
      continue;
    }
    if (css[cursor] === "@") {
      const name = consumeCssName(css, cursor + 1);
      if (name.value.toLowerCase() === "import") {
        let end = name.end;
        let depth = 0;
        while (end < css.length) {
          if (css.startsWith("/*", end)) end = consumeCssComment(css, end);
          else if (css[end] === "\"" || css[end] === "'") end = consumeCssString(css, end);
          else if (css[end] === "(") { depth++; end++; }
          else if (css[end] === ")") { depth = Math.max(0, depth - 1); end++; }
          else if (css[end] === ";" && depth === 0) { end++; break; }
          else end++;
        }
        tokens.push({ kind: "import", start: cursor, end, value: css.slice(cursor, end) });
        cursor = end;
        continue;
      }
    }
    if (css[cursor] === "\\" || /[a-z_-]/i.test(css[cursor]) || css.charCodeAt(cursor) >= 0x80) {
      const name = consumeCssName(css, cursor);
      if (name.value.toLowerCase() === "url" && css[name.end] === "(") {
        let end = name.end + 1;
        while (end < css.length) {
          if (css.startsWith("/*", end)) end = consumeCssComment(css, end);
          else if (css[end] === "\"" || css[end] === "'") end = consumeCssString(css, end);
          else if (css[end] === "\\") end = consumeCssEscape(css, end).end;
          else if (css[end++] === ")") break;
        }
        const innerEnd = css[end - 1] === ")" ? end - 1 : end;
        tokens.push({
          kind: "url",
          start: cursor,
          end,
          value: decodeCssUrlComponent(css.slice(name.end + 1, innerEnd)),
        });
        cursor = end;
        continue;
      }
      if (css[name.end] === "(") {
        const functionName = name.value.toLowerCase();
        const isImageSource = functionStack.some(
          (context) => context === "image-set" || context === "-webkit-image-set",
        );
        if (isImageSource
          && (functionName === "var" || functionName === "env" || functionName === "if")) {
          const end = consumeCssFunction(css, name.end);
          tokens.push({
            kind: "substitution",
            start: cursor,
            end,
            value: css.slice(cursor, end),
          });
          cursor = end;
          continue;
        }
        functionStack.push(functionName);
        cursor = name.end + 1;
        continue;
      }
      cursor = Math.max(name.end, cursor + 1);
      continue;
    }
    if (css[cursor] === "(") {
      functionStack.push("");
      cursor++;
      continue;
    }
    if (css[cursor] === ")") {
      functionStack.pop();
      cursor++;
      continue;
    }
    cursor++;
  }
  return tokens;
}
