import * as fs from 'fs';

type ZipEntry = {
  name: string;
  crc32: number;
  size: number;
  localHeaderOffset: number;
};

const ZIP32_MAX_UINT16 = 0xffff;
const ZIP32_MAX_UINT32 = 0xffffffff;

export type StoredZipFile = {
  name: string;
  data: Buffer;
};

function writeUint32LE(buffer: Buffer, value: number, offset: number): number {
  const normalized = value >>> 0;
  buffer[offset] = normalized & 0xff;
  buffer[offset + 1] = (normalized >>> 8) & 0xff;
  buffer[offset + 2] = (normalized >>> 16) & 0xff;
  buffer[offset + 3] = (normalized >>> 24) & 0xff;
  return offset + 4;
}

const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let crc = i;
    for (let j = 0; j < 8; j += 1) {
      crc = (crc & 1) !== 0 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
    }
    table[i] = crc >>> 0;
  }
  return table;
})();

function crc32(data: Buffer): number {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = CRC32_TABLE[(crc ^ byte) & 0xff]! ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function createLocalFileHeader(fileName: Buffer, fileCrc32: number, fileSize: number): Buffer {
  const local = Buffer.alloc(30 + fileName.length);
  let cursor = 0;
  writeUint32LE(local, 0x04034b50, cursor);
  cursor += 4;
  local.writeUInt16LE(20, cursor);
  cursor += 2;
  local.writeUInt16LE(0, cursor);
  cursor += 2;
  local.writeUInt16LE(0, cursor);
  cursor += 2;
  local.writeUInt16LE(0, cursor);
  cursor += 2;
  local.writeUInt16LE(0, cursor);
  cursor += 2;
  writeUint32LE(local, fileCrc32, cursor);
  cursor += 4;
  writeUint32LE(local, fileSize, cursor);
  cursor += 4;
  writeUint32LE(local, fileSize, cursor);
  cursor += 4;
  local.writeUInt16LE(fileName.length, cursor);
  cursor += 2;
  local.writeUInt16LE(0, cursor);
  cursor += 2;
  fileName.copy(local, cursor);
  return local;
}

function createCentralDirectoryHeader(entry: ZipEntry): Buffer {
  const fileName = Buffer.from(entry.name, 'utf8');
  const central = Buffer.alloc(46 + fileName.length);
  let cursor = 0;
  writeUint32LE(central, 0x02014b50, cursor);
  cursor += 4;
  central.writeUInt16LE(20, cursor);
  cursor += 2;
  central.writeUInt16LE(20, cursor);
  cursor += 2;
  central.writeUInt16LE(0, cursor);
  cursor += 2;
  central.writeUInt16LE(0, cursor);
  cursor += 2;
  central.writeUInt16LE(0, cursor);
  cursor += 2;
  central.writeUInt16LE(0, cursor);
  cursor += 2;
  writeUint32LE(central, entry.crc32, cursor);
  cursor += 4;
  writeUint32LE(central, entry.size, cursor);
  cursor += 4;
  writeUint32LE(central, entry.size, cursor);
  cursor += 4;
  central.writeUInt16LE(fileName.length, cursor);
  cursor += 2;
  central.writeUInt16LE(0, cursor);
  cursor += 2;
  central.writeUInt16LE(0, cursor);
  cursor += 2;
  central.writeUInt16LE(0, cursor);
  cursor += 2;
  central.writeUInt16LE(0, cursor);
  cursor += 2;
  writeUint32LE(central, 0, cursor);
  cursor += 4;
  writeUint32LE(central, entry.localHeaderOffset, cursor);
  cursor += 4;
  fileName.copy(central, cursor);
  return central;
}

function createEndOfCentralDirectory(
  entriesLength: number,
  centralSize: number,
  centralStart: number,
): Buffer {
  if (
    entriesLength > ZIP32_MAX_UINT16 ||
    centralSize > ZIP32_MAX_UINT32 ||
    centralStart > ZIP32_MAX_UINT32
  ) {
    throw new RangeError('Archive exceeds ZIP32 limits (Zip64 not implemented)');
  }

  const end = Buffer.alloc(22);
  let cursor = 0;
  writeUint32LE(end, 0x06054b50, cursor);
  cursor += 4;
  end.writeUInt16LE(0, cursor);
  cursor += 2;
  end.writeUInt16LE(0, cursor);
  cursor += 2;
  end.writeUInt16LE(entriesLength, cursor);
  cursor += 2;
  end.writeUInt16LE(entriesLength, cursor);
  cursor += 2;
  writeUint32LE(end, centralSize, cursor);
  cursor += 4;
  writeUint32LE(end, centralStart, cursor);
  cursor += 4;
  end.writeUInt16LE(0, cursor);
  return end;
}

function writeBuffer(fd: number, buffer: Buffer): void {
  let written = 0;
  while (written < buffer.length) {
    written += fs.writeSync(fd, buffer, written, buffer.length - written);
  }
}

export function writeStoredZip(
  outputPath: string,
  files: Iterable<StoredZipFile>,
): { entryCount: number } {
  const entries: ZipEntry[] = [];
  let offset = 0;
  const fd = fs.openSync(outputPath, 'w');

  try {
    for (const file of files) {
      const fileName = Buffer.from(file.name, 'utf8');
      const fileSize = file.data.length;
      if (fileName.length > ZIP32_MAX_UINT16) {
        throw new RangeError(`ZIP entry name too long: ${file.name}`);
      }
      if (fileSize > ZIP32_MAX_UINT32) {
        throw new RangeError(`ZIP entry too large for ZIP32: ${file.name}`);
      }
      if (offset > ZIP32_MAX_UINT32) {
        throw new RangeError('Archive exceeds ZIP32 limits (Zip64 not implemented)');
      }
      const fileCrc32 = crc32(file.data);
      const localHeader = createLocalFileHeader(fileName, fileCrc32, fileSize);
      const nextOffset = offset + localHeader.length + fileSize;
      if (nextOffset > ZIP32_MAX_UINT32) {
        throw new RangeError('Archive exceeds ZIP32 limits (Zip64 not implemented)');
      }
      writeBuffer(fd, localHeader);
      writeBuffer(fd, file.data);
      entries.push({
        name: file.name,
        crc32: fileCrc32,
        size: fileSize,
        localHeaderOffset: offset,
      });
      if (nextOffset > ZIP32_MAX_UINT32) {
        throw new RangeError('Archive exceeds ZIP32 limits (Zip64 not implemented)');
      }
      offset = nextOffset;
    }

    const centralStart = offset;
    if (centralStart > ZIP32_MAX_UINT32) {
      throw new RangeError('Archive exceeds ZIP32 limits (Zip64 not implemented)');
    }
    for (const entry of entries) {
      const centralHeader = createCentralDirectoryHeader(entry);
      writeBuffer(fd, centralHeader);
      offset += centralHeader.length;
    }

    const centralSize = offset - centralStart;
    writeBuffer(fd, createEndOfCentralDirectory(entries.length, centralSize, centralStart));
  } catch (error) {
    fs.closeSync(fd);
    fs.rmSync(outputPath, { force: true });
    throw error;
  }

  fs.closeSync(fd);
  return { entryCount: entries.length };
}
