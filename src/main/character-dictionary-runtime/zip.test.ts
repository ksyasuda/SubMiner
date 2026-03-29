import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { buildDictionaryZip } from './zip';
import type { CharacterDictionaryTermEntry } from './types';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-character-zip-'));
}

function cleanupDir(dirPath: string): void {
  fs.rmSync(dirPath, { recursive: true, force: true });
}

function readStoredZipEntries(zipPath: string): Map<string, Buffer> {
  const archive = fs.readFileSync(zipPath);
  const entries = new Map<string, Buffer>();
  let cursor = 0;

  while (cursor + 4 <= archive.length) {
    const signature = archive.readUInt32LE(cursor);
    if (signature === 0x02014b50 || signature === 0x06054b50) {
      break;
    }
    assert.equal(signature, 0x04034b50, `unexpected local file header at offset ${cursor}`);

    const compressedSize = archive.readUInt32LE(cursor + 18);
    const fileNameLength = archive.readUInt16LE(cursor + 26);
    const extraLength = archive.readUInt16LE(cursor + 28);
    const fileNameStart = cursor + 30;
    const dataStart = fileNameStart + fileNameLength + extraLength;
    const fileName = archive.subarray(fileNameStart, fileNameStart + fileNameLength).toString(
      'utf8',
    );
    const data = archive.subarray(dataStart, dataStart + compressedSize);
    entries.set(fileName, Buffer.from(data));
    cursor = dataStart + compressedSize;
  }

  return entries;
}

test('buildDictionaryZip writes a valid stored zip without fs.writeFileSync', () => {
  const tempDir = makeTempDir();
  const outputPath = path.join(tempDir, 'dictionary.zip');
  const termEntries: CharacterDictionaryTermEntry[] = [
    ['アルファ', 'あるふぁ', '', '', 0, ['Alpha entry'], 0, 'name'],
  ];
  const originalWriteFileSync = fs.writeFileSync;
  const originalBufferConcat = Buffer.concat;

  try {
    fs.writeFileSync = ((..._args: unknown[]) => {
      throw new Error('buildDictionaryZip should not call fs.writeFileSync');
    }) as typeof fs.writeFileSync;

    Buffer.concat = ((...args: Parameters<typeof Buffer.concat>) => {
      throw new Error(`buildDictionaryZip should not Buffer.concat the full archive (${args[0].length} chunks)`);
    }) as typeof Buffer.concat;

    const result = buildDictionaryZip(
      outputPath,
      'Dictionary Title',
      'Dictionary Description',
      '2026-03-27',
      termEntries,
      [{ path: 'images/alpha.bin', dataBase64: Buffer.from([1, 2, 3]).toString('base64') }],
    );

    assert.equal(result.zipPath, outputPath);
    assert.equal(result.entryCount, 1);

    const entries = readStoredZipEntries(outputPath);
    assert.deepEqual([...entries.keys()].sort(), [
      'images/alpha.bin',
      'index.json',
      'tag_bank_1.json',
      'term_bank_1.json',
    ]);

    const indexJson = JSON.parse(entries.get('index.json')!.toString('utf8')) as {
      title: string;
      description: string;
      revision: string;
      format: number;
    };
    assert.equal(indexJson.title, 'Dictionary Title');
    assert.equal(indexJson.description, 'Dictionary Description');
    assert.equal(indexJson.revision, '2026-03-27');
    assert.equal(indexJson.format, 3);

    const termBank = JSON.parse(entries.get('term_bank_1.json')!.toString('utf8')) as
      CharacterDictionaryTermEntry[];
    assert.equal(termBank.length, 1);
    assert.equal(termBank[0]?.[0], 'アルファ');
    assert.deepEqual(entries.get('images/alpha.bin'), Buffer.from([1, 2, 3]));
  } finally {
    fs.writeFileSync = originalWriteFileSync;
    Buffer.concat = originalBufferConcat;
    cleanupDir(tempDir);
  }
});
