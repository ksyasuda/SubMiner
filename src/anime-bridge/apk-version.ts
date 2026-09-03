import { open, type FileHandle } from 'node:fs/promises';
import { inflateRawSync } from 'node:zlib';

const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;
const CENTRAL_FILE_HEADER_SIGNATURE = 0x02014b50;
const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;
const ANDROID_XML_TYPE = 0x0003;
const STRING_POOL_TYPE = 0x0001;
const START_ELEMENT_TYPE = 0x0102;
const UTF8_STRING_POOL_FLAG = 0x00000100;
const NO_STRING = 0xffffffff;
const TYPE_INT_DEC = 0x10;
const TYPE_INT_HEX = 0x11;
const MANIFEST_ENTRY = 'AndroidManifest.xml';
const MAX_ZIP_TAIL_BYTES = 65_535 + 22;
const MAX_CENTRAL_DIRECTORY_BYTES = 16 * 1024 * 1024;
const MAX_MANIFEST_BYTES = 1024 * 1024;

interface ZipEntryLocation {
  compressionMethod: number;
  compressedSize: number;
  uncompressedSize: number;
  localHeaderOffset: number;
}

async function readExactly(handle: FileHandle, length: number, position: number): Promise<Buffer> {
  const buffer = Buffer.alloc(length);
  const { bytesRead } = await handle.read(buffer, 0, length, position);
  if (bytesRead !== length) throw new Error('Unexpected end of APK.');
  return buffer;
}

function findEndOfCentralDirectory(tail: Buffer): number | null {
  for (let offset = tail.length - 22; offset >= 0; offset -= 1) {
    if (tail.readUInt32LE(offset) !== END_OF_CENTRAL_DIRECTORY_SIGNATURE) continue;
    const commentLength = tail.readUInt16LE(offset + 20);
    if (offset + 22 + commentLength === tail.length) return offset;
  }
  return null;
}

function findZipEntry(
  central: Buffer,
  entryCount: number,
  wantedName: string,
): ZipEntryLocation | null {
  let offset = 0;
  for (let index = 0; index < entryCount; index += 1) {
    if (offset + 46 > central.length) return null;
    if (central.readUInt32LE(offset) !== CENTRAL_FILE_HEADER_SIGNATURE) return null;

    const flags = central.readUInt16LE(offset + 8);
    const compressionMethod = central.readUInt16LE(offset + 10);
    const compressedSize = central.readUInt32LE(offset + 20);
    const uncompressedSize = central.readUInt32LE(offset + 24);
    const nameLength = central.readUInt16LE(offset + 28);
    const extraLength = central.readUInt16LE(offset + 30);
    const commentLength = central.readUInt16LE(offset + 32);
    const recordLength = 46 + nameLength + extraLength + commentLength;
    if (offset + recordLength > central.length) return null;

    const name = central.subarray(offset + 46, offset + 46 + nameLength).toString('utf8');
    if (name === wantedName) {
      if ((flags & 0x0001) !== 0) return null;
      if (compressedSize > MAX_MANIFEST_BYTES || uncompressedSize > MAX_MANIFEST_BYTES) return null;
      return {
        compressionMethod,
        compressedSize,
        uncompressedSize,
        localHeaderOffset: central.readUInt32LE(offset + 42),
      };
    }
    offset += recordLength;
  }
  return null;
}

async function readZipEntry(apkPath: string, wantedName: string): Promise<Buffer | null> {
  let handle: FileHandle | null = null;
  try {
    handle = await open(apkPath, 'r');
    const fileSize = (await handle.stat()).size;
    const tailSize = Math.min(fileSize, MAX_ZIP_TAIL_BYTES);
    if (tailSize < 22) return null;
    const tail = await readExactly(handle, tailSize, fileSize - tailSize);
    const endOffset = findEndOfCentralDirectory(tail);
    if (endOffset === null) return null;

    const diskNumber = tail.readUInt16LE(endOffset + 4);
    const centralDisk = tail.readUInt16LE(endOffset + 6);
    const diskEntryCount = tail.readUInt16LE(endOffset + 8);
    const entryCount = tail.readUInt16LE(endOffset + 10);
    const centralSize = tail.readUInt32LE(endOffset + 12);
    const centralOffset = tail.readUInt32LE(endOffset + 16);
    if (
      diskNumber !== 0 ||
      centralDisk !== 0 ||
      diskEntryCount !== entryCount ||
      entryCount === 0 ||
      centralSize === 0 ||
      centralSize > MAX_CENTRAL_DIRECTORY_BYTES ||
      centralOffset + centralSize > fileSize
    ) {
      return null;
    }

    const central = await readExactly(handle, centralSize, centralOffset);
    const entry = findZipEntry(central, entryCount, wantedName);
    if (!entry || entry.localHeaderOffset + 30 > centralOffset) return null;

    const localHeader = await readExactly(handle, 30, entry.localHeaderOffset);
    if (localHeader.readUInt32LE(0) !== LOCAL_FILE_HEADER_SIGNATURE) return null;
    const nameLength = localHeader.readUInt16LE(26);
    const extraLength = localHeader.readUInt16LE(28);
    const dataOffset = entry.localHeaderOffset + 30 + nameLength + extraLength;
    if (dataOffset + entry.compressedSize > centralOffset) return null;
    const compressed = await readExactly(handle, entry.compressedSize, dataOffset);

    let data: Buffer;
    if (entry.compressionMethod === 0) {
      data = compressed;
    } else if (entry.compressionMethod === 8) {
      data = inflateRawSync(compressed, { maxOutputLength: MAX_MANIFEST_BYTES });
    } else {
      return null;
    }
    return data.length === entry.uncompressedSize ? data : null;
  } catch {
    return null;
  } finally {
    await handle?.close().catch(() => undefined);
  }
}

function readUtf8Length(buffer: Buffer, offset: number): { length: number; next: number } | null {
  if (offset >= buffer.length) return null;
  const first = buffer[offset]!;
  if ((first & 0x80) === 0) return { length: first, next: offset + 1 };
  if (offset + 1 >= buffer.length) return null;
  return { length: ((first & 0x7f) << 8) | buffer[offset + 1]!, next: offset + 2 };
}

function readUtf16Length(buffer: Buffer, offset: number): { length: number; next: number } | null {
  if (offset + 2 > buffer.length) return null;
  const first = buffer.readUInt16LE(offset);
  if ((first & 0x8000) === 0) return { length: first, next: offset + 2 };
  if (offset + 4 > buffer.length) return null;
  return {
    length: ((first & 0x7fff) << 16) | buffer.readUInt16LE(offset + 2),
    next: offset + 4,
  };
}

interface AndroidStringPool {
  stringAt: (index: number) => string | null;
}

function parseStringPool(
  buffer: Buffer,
  chunkOffset: number,
  chunkSize: number,
): AndroidStringPool | null {
  const headerSize = buffer.readUInt16LE(chunkOffset + 2);
  if (headerSize < 28 || chunkOffset + chunkSize > buffer.length) return null;
  const stringCount = buffer.readUInt32LE(chunkOffset + 8);
  const flags = buffer.readUInt32LE(chunkOffset + 16);
  const stringsStart = buffer.readUInt32LE(chunkOffset + 20);
  if (headerSize + stringCount * 4 > chunkSize || stringsStart >= chunkSize) return null;

  return {
    stringAt(index) {
      if (index === NO_STRING || index >= stringCount) return null;
      const relativeOffset = buffer.readUInt32LE(chunkOffset + headerSize + index * 4);
      let stringOffset = chunkOffset + stringsStart + relativeOffset;
      const chunkEnd = chunkOffset + chunkSize;
      if (stringOffset >= chunkEnd) return null;

      if ((flags & UTF8_STRING_POOL_FLAG) !== 0) {
        const utf16Length = readUtf8Length(buffer, stringOffset);
        if (!utf16Length) return null;
        const byteLength = readUtf8Length(buffer, utf16Length.next);
        if (!byteLength || byteLength.next + byteLength.length > chunkEnd) return null;
        return buffer.toString('utf8', byteLength.next, byteLength.next + byteLength.length);
      }

      const length = readUtf16Length(buffer, stringOffset);
      if (!length) return null;
      stringOffset = length.next;
      const byteLength = length.length * 2;
      if (stringOffset + byteLength > chunkEnd) return null;
      return buffer.toString('utf16le', stringOffset, stringOffset + byteLength);
    },
  };
}

/** Read Android's numeric version code from a binary AndroidManifest.xml. */
export function parseAndroidManifestVersionCode(buffer: Buffer): number | null {
  try {
    if (buffer.length < 8 || buffer.readUInt16LE(0) !== ANDROID_XML_TYPE) return null;
    const documentSize = buffer.readUInt32LE(4);
    if (documentSize > buffer.length) return null;

    let strings: AndroidStringPool | null = null;
    let offset = buffer.readUInt16LE(2);
    while (offset + 8 <= documentSize) {
      const chunkType = buffer.readUInt16LE(offset);
      const headerSize = buffer.readUInt16LE(offset + 2);
      const chunkSize = buffer.readUInt32LE(offset + 4);
      if (headerSize < 8 || chunkSize < headerSize || offset + chunkSize > documentSize)
        return null;

      if (chunkType === STRING_POOL_TYPE) {
        strings = parseStringPool(buffer, offset, chunkSize);
      } else if (chunkType === START_ELEMENT_TYPE && strings && headerSize >= 16) {
        const elementName = strings.stringAt(buffer.readUInt32LE(offset + 20));
        if (elementName === 'manifest') {
          const attributeStart = buffer.readUInt16LE(offset + 24);
          const attributeSize = buffer.readUInt16LE(offset + 26);
          const attributeCount = buffer.readUInt16LE(offset + 28);
          const attributesOffset = offset + 16 + attributeStart;
          if (
            attributeSize < 20 ||
            attributesOffset + attributeSize * attributeCount > offset + chunkSize
          ) {
            return null;
          }

          for (let index = 0; index < attributeCount; index += 1) {
            const attributeOffset = attributesOffset + index * attributeSize;
            const name = strings.stringAt(buffer.readUInt32LE(attributeOffset + 4));
            if (name !== 'versionCode') continue;
            const valueType = buffer[attributeOffset + 15];
            if (valueType === TYPE_INT_DEC || valueType === TYPE_INT_HEX) {
              return buffer.readUInt32LE(attributeOffset + 16);
            }
            const rawValue = strings.stringAt(buffer.readUInt32LE(attributeOffset + 8));
            if (rawValue === null) return null;
            const parsed = Number(rawValue);
            return Number.isSafeInteger(parsed) && parsed >= 0 ? parsed : null;
          }
          return null;
        }
      }
      offset += chunkSize;
    }
    return null;
  } catch {
    return null;
  }
}

/** Read the installed version without extracting the APK or running Android tooling. */
export async function readApkVersionCode(apkPath: string): Promise<number | null> {
  const manifest = await readZipEntry(apkPath, MANIFEST_ENTRY);
  return manifest ? parseAndroidManifestVersionCode(manifest) : null;
}
