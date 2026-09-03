import assert from 'node:assert/strict';
import test from 'node:test';
import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { deflateRawSync } from 'node:zlib';
import { parseAndroidManifestVersionCode, readApkVersionCode } from './apk-version';

const NO_STRING = 0xffffffff;

function writeChunkHeader(buffer: Buffer, type: number, headerSize: number, size: number): void {
  buffer.writeUInt16LE(type, 0);
  buffer.writeUInt16LE(headerSize, 2);
  buffer.writeUInt32LE(size, 4);
}

function makeStringPool(strings: string[]): Buffer {
  const encoded = strings.map((value) => {
    const bytes = Buffer.from(value, 'utf8');
    return Buffer.concat([Buffer.from([value.length, bytes.length]), bytes, Buffer.from([0])]);
  });
  const offsets = encoded.map((_value, index) =>
    encoded.slice(0, index).reduce((total, value) => total + value.length, 0),
  );
  const dataLength = encoded.reduce((total, value) => total + value.length, 0);
  const paddedDataLength = Math.ceil(dataLength / 4) * 4;
  const headerSize = 28;
  const stringsStart = headerSize + strings.length * 4;
  const chunk = Buffer.alloc(stringsStart + paddedDataLength);
  writeChunkHeader(chunk, 0x0001, headerSize, chunk.length);
  chunk.writeUInt32LE(strings.length, 8);
  chunk.writeUInt32LE(0x00000100, 16);
  chunk.writeUInt32LE(stringsStart, 20);
  offsets.forEach((offset, index) => chunk.writeUInt32LE(offset, headerSize + index * 4));
  Buffer.concat(encoded).copy(chunk, stringsStart);
  return chunk;
}

function makeBinaryManifest(versionCode: number): Buffer {
  const stringPool = makeStringPool(['manifest', 'versionCode']);
  const startElement = Buffer.alloc(56);
  writeChunkHeader(startElement, 0x0102, 16, startElement.length);
  startElement.writeUInt32LE(NO_STRING, 12);
  startElement.writeUInt32LE(NO_STRING, 16);
  startElement.writeUInt32LE(0, 20);
  startElement.writeUInt16LE(20, 24);
  startElement.writeUInt16LE(20, 26);
  startElement.writeUInt16LE(1, 28);
  const attributeOffset = 36;
  startElement.writeUInt32LE(NO_STRING, attributeOffset);
  startElement.writeUInt32LE(1, attributeOffset + 4);
  startElement.writeUInt32LE(NO_STRING, attributeOffset + 8);
  startElement.writeUInt16LE(8, attributeOffset + 12);
  startElement[attributeOffset + 15] = 0x10;
  startElement.writeUInt32LE(versionCode, attributeOffset + 16);

  const document = Buffer.alloc(8);
  writeChunkHeader(document, 0x0003, 8, document.length + stringPool.length + startElement.length);
  return Buffer.concat([document, stringPool, startElement]);
}

function makeDeflatedZip(name: string, data: Buffer): Buffer {
  const fileName = Buffer.from(name, 'utf8');
  const compressed = deflateRawSync(data);
  const local = Buffer.alloc(30 + fileName.length);
  local.writeUInt32LE(0x04034b50, 0);
  local.writeUInt16LE(20, 4);
  local.writeUInt16LE(8, 8);
  local.writeUInt32LE(compressed.length, 18);
  local.writeUInt32LE(data.length, 22);
  local.writeUInt16LE(fileName.length, 26);
  fileName.copy(local, 30);

  const central = Buffer.alloc(46 + fileName.length);
  central.writeUInt32LE(0x02014b50, 0);
  central.writeUInt16LE(20, 4);
  central.writeUInt16LE(20, 6);
  central.writeUInt16LE(8, 10);
  central.writeUInt32LE(compressed.length, 20);
  central.writeUInt32LE(data.length, 24);
  central.writeUInt16LE(fileName.length, 28);
  fileName.copy(central, 46);

  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(1, 8);
  end.writeUInt16LE(1, 10);
  end.writeUInt32LE(central.length, 12);
  end.writeUInt32LE(local.length + compressed.length, 16);
  return Buffer.concat([local, compressed, central, end]);
}

test('parseAndroidManifestVersionCode reads the typed manifest attribute', () => {
  assert.equal(parseAndroidManifestVersionCode(makeBinaryManifest(42)), 42);
  assert.equal(parseAndroidManifestVersionCode(Buffer.from('plain xml')), null);
});

test('readApkVersionCode reads a deflated AndroidManifest.xml from an APK', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'subminer-apk-version-'));
  const apkPath = path.join(directory, 'extension.apk');
  await writeFile(apkPath, makeDeflatedZip('AndroidManifest.xml', makeBinaryManifest(730)));

  assert.equal(await readApkVersionCode(apkPath), 730);
});

test('readApkVersionCode returns null for a malformed APK', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'subminer-apk-version-'));
  const apkPath = path.join(directory, 'broken.apk');
  await writeFile(apkPath, 'not a zip');

  assert.equal(await readApkVersionCode(apkPath), null);
});
