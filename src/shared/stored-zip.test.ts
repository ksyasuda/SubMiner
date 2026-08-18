import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { readStoredZipFirstFile, writeStoredZip, writeStoredZipAsync } from './stored-zip';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-stored-zip-'));
}

function readEntries(zipPath: string): Map<string, Buffer> {
  const archive = fs.readFileSync(zipPath);
  const entries = new Map<string, Buffer>();
  let cursor = 0;

  while (cursor + 4 <= archive.length) {
    const signature = archive.readUInt32LE(cursor);
    if (signature === 0x02014b50 || signature === 0x06054b50) {
      break;
    }
    assert.equal(signature, 0x04034b50, `unexpected local file header at offset ${cursor}`);
    const size = archive.readUInt32LE(cursor + 18);
    const nameLength = archive.readUInt16LE(cursor + 26);
    const extraLength = archive.readUInt16LE(cursor + 28);
    const nameStart = cursor + 30;
    const dataStart = nameStart + nameLength + extraLength;
    entries.set(
      archive.subarray(nameStart, nameStart + nameLength).toString('utf8'),
      Buffer.from(archive.subarray(dataStart, dataStart + size)),
    );
    cursor = dataStart + size;
  }

  return entries;
}

// The async writer yields on a byte budget between entries, so an entry that is itself larger than
// that budget is the case where the accounting could drift: the offsets, CRCs, and central
// directory all have to come out identical to the synchronous writer.
test('writeStoredZipAsync writes a correct archive when one entry exceeds the yield budget', async () => {
  const dir = makeTempDir();
  try {
    // Comfortably past the writer's 8MB yield budget.
    const oversized = Buffer.alloc(10 * 1024 * 1024);
    for (let i = 0; i < oversized.length; i += 1) {
      oversized[i] = i % 251;
    }
    const files = [
      { name: 'index.json', data: Buffer.from('{"revision":"rev-1"}', 'utf8') },
      { name: 'big.bin', data: oversized },
      { name: 'after.txt', data: Buffer.from('written after the oversized entry', 'utf8') },
    ];

    const asyncPath = path.join(dir, 'async.zip');
    const syncPath = path.join(dir, 'sync.zip');
    const asyncResult = await writeStoredZipAsync(asyncPath, files);
    const syncResult = writeStoredZip(syncPath, files);

    assert.equal(asyncResult.entryCount, 3);
    assert.deepEqual(asyncResult, syncResult);
    // Byte-identical to the synchronous writer: yielding mid-archive changed no offset or CRC.
    assert.ok(fs.readFileSync(asyncPath).equals(fs.readFileSync(syncPath)));

    const entries = readEntries(asyncPath);
    assert.deepEqual([...entries.keys()], ['index.json', 'big.bin', 'after.txt']);
    assert.ok(entries.get('big.bin')!.equals(oversized));
    assert.equal(entries.get('after.txt')!.toString('utf8'), 'written after the oversized entry');
    // The trailing records still parse, which is what proves the archive is complete.
    assert.equal(readStoredZipFirstFile(asyncPath)?.name, 'index.json');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
