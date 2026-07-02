const EDL_URI_PREFIX = 'edl://';

const BYTE_COMMA = ','.charCodeAt(0);
const BYTE_CR = '\r'.charCodeAt(0);
const BYTE_EQUALS = '='.charCodeAt(0);
const BYTE_EXCLAMATION = '!'.charCodeAt(0);
const BYTE_LF = '\n'.charCodeAt(0);
const BYTE_PERCENT = '%'.charCodeAt(0);
const BYTE_SEMICOLON = ';'.charCodeAt(0);

function isDigitByte(value: number | undefined): value is number {
  return value !== undefined && value >= 48 && value <= 57;
}

function isEntrySeparator(value: number | undefined): boolean {
  return value === BYTE_SEMICOLON || value === BYTE_LF || value === BYTE_CR;
}

function isParamSeparator(value: number | undefined): boolean {
  return value === BYTE_COMMA || isEntrySeparator(value);
}

function decodeBytes(buffer: Buffer, start: number, end: number): string {
  return buffer.subarray(start, end).toString('utf8');
}

function isHttpUrl(value: string): boolean {
  return /^https?:\/\//i.test(value);
}

function toEdlDataBuffer(source: string): Buffer {
  const data = source.startsWith(EDL_URI_PREFIX) ? source.slice(EDL_URI_PREFIX.length) : source;
  return Buffer.from(data, 'utf8');
}

function parseLengthGuardedValue(
  buffer: Buffer,
  position: number,
): { value: string; end: number } | null {
  if (buffer[position] !== BYTE_PERCENT) {
    return null;
  }

  let cursor = position + 1;
  if (!isDigitByte(buffer[cursor])) {
    return null;
  }

  let byteLength = 0;
  while (true) {
    const digit = buffer[cursor];
    if (!isDigitByte(digit)) {
      break;
    }
    byteLength = byteLength * 10 + (digit - 48);
    cursor += 1;
  }

  if (buffer[cursor] !== BYTE_PERCENT) {
    return null;
  }

  const valueStart = cursor + 1;
  const valueEnd = valueStart + byteLength;
  if (valueEnd > buffer.length) {
    return null;
  }

  return {
    value: decodeBytes(buffer, valueStart, valueEnd),
    end: valueEnd,
  };
}

function skipEntrySeparators(buffer: Buffer, position: number): number {
  let cursor = position;
  while (cursor < buffer.length && isEntrySeparator(buffer[cursor])) {
    cursor += 1;
  }
  return cursor;
}

function skipEntry(buffer: Buffer, position: number): number {
  let cursor = position;
  while (cursor < buffer.length) {
    const guardedValue = parseLengthGuardedValue(buffer, cursor);
    if (guardedValue) {
      cursor = guardedValue.end;
      continue;
    }
    if (isEntrySeparator(buffer[cursor])) {
      break;
    }
    cursor += 1;
  }
  return cursor;
}

function parseRawValue(buffer: Buffer, position: number): { value: string; end: number } {
  let cursor = position;
  while (
    cursor < buffer.length &&
    !isParamSeparator(buffer[cursor]) &&
    buffer[cursor] !== BYTE_EXCLAMATION
  ) {
    cursor += 1;
  }
  return {
    value: decodeBytes(buffer, position, cursor),
    end: cursor,
  };
}

function parseParamValue(buffer: Buffer, position: number): { value: string; end: number } {
  return parseLengthGuardedValue(buffer, position) ?? parseRawValue(buffer, position);
}

function parseOptionalParamName(
  buffer: Buffer,
  position: number,
): { name: string | null; valueStart: number } {
  let cursor = position;
  while (
    cursor < buffer.length &&
    !isParamSeparator(buffer[cursor]) &&
    buffer[cursor] !== BYTE_PERCENT &&
    buffer[cursor] !== BYTE_EXCLAMATION
  ) {
    if (buffer[cursor] === BYTE_EQUALS) {
      return {
        name: decodeBytes(buffer, position, cursor),
        valueStart: cursor + 1,
      };
    }
    cursor += 1;
  }

  return { name: null, valueStart: position };
}

function parseSegmentEntry(buffer: Buffer, position: number): { urls: string[]; end: number } {
  const urls: string[] = [];
  let cursor = position;
  let unnamedParamIndex = 0;

  while (cursor < buffer.length && !isEntrySeparator(buffer[cursor])) {
    const { name, valueStart } = parseOptionalParamName(buffer, cursor);
    const value = parseParamValue(buffer, valueStart);
    const lowerName = name?.toLowerCase() ?? null;
    const isFileParam = lowerName === 'file' || (lowerName === null && unnamedParamIndex === 0);

    if (isFileParam && isHttpUrl(value.value)) {
      urls.push(value.value);
    }

    if (lowerName === null) {
      unnamedParamIndex += 1;
    }

    cursor = value.end;
    if (buffer[cursor] === BYTE_COMMA) {
      cursor += 1;
      continue;
    }
    if (!isEntrySeparator(buffer[cursor])) {
      cursor = skipEntry(buffer, cursor);
    }
  }

  return { urls, end: cursor };
}

export function extractFileUrlsFromMpvEdlSource(source: string): string[] {
  const buffer = toEdlDataBuffer(source);
  const urls: string[] = [];
  let cursor = 0;

  while (cursor < buffer.length) {
    cursor = skipEntrySeparators(buffer, cursor);
    if (cursor >= buffer.length) {
      break;
    }

    if (buffer[cursor] === BYTE_EXCLAMATION) {
      cursor = skipEntry(buffer, cursor);
      continue;
    }

    const segment = parseSegmentEntry(buffer, cursor);
    urls.push(...segment.urls);
    cursor = segment.end;
  }

  return urls;
}
