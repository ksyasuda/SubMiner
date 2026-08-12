import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { buildDictionaryZip, readDictionaryZipRevision } from './zip';
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
    const fileName = archive
      .subarray(fileNameStart, fileNameStart + fileNameLength)
      .toString('utf8');
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
      throw new Error(
        `buildDictionaryZip should not Buffer.concat the full archive (${args[0].length} chunks)`,
      );
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

    const termBank = JSON.parse(
      entries.get('term_bank_1.json')!.toString('utf8'),
    ) as CharacterDictionaryTermEntry[];
    assert.equal(termBank.length, 1);
    assert.equal(termBank[0]?.[0], 'アルファ');
    assert.deepEqual(entries.get('images/alpha.bin'), Buffer.from([1, 2, 3]));
  } finally {
    fs.writeFileSync = originalWriteFileSync;
    Buffer.concat = originalBufferConcat;
    cleanupDir(tempDir);
  }
});

test('readDictionaryZipRevision reads the built revision and rejects foreign archives', () => {
  const dir = makeTempDir();
  try {
    const zipPath = path.join(dir, 'merged.zip');
    buildDictionaryZip(
      zipPath,
      'SubMiner Character Dictionary',
      'Character names',
      'rev-42',
      [
        {
          term: 'ルフィ',
          reading: 'ルフィ',
          role: 'main',
          glossary: [],
        } as unknown as CharacterDictionaryTermEntry,
      ],
      [],
    );

    assert.equal(readDictionaryZipRevision(zipPath), 'rev-42');
    assert.equal(readDictionaryZipRevision(path.join(dir, 'missing.zip')), null);

    const archive = fs.readFileSync(zipPath);
    const truncatedPath = path.join(dir, 'truncated.zip');
    fs.writeFileSync(truncatedPath, archive.subarray(0, 40));
    assert.equal(readDictionaryZipRevision(truncatedPath), null);

    // An archive cut short after index.json still holds a readable revision, but importing it
    // would hand Yomitan a half-written file: the missing end-of-central-directory record has to
    // reject it. One byte off the end is enough to make the record incomplete.
    for (const missingBytes of [1, 22, archive.length - 200]) {
      const cutPath = path.join(dir, `cut-${missingBytes}.zip`);
      fs.writeFileSync(cutPath, archive.subarray(0, archive.length - missingBytes));
      assert.equal(readDictionaryZipRevision(cutPath), null, `cut of ${missingBytes} bytes`);
    }

    // Same size, corrupt directory: a record overwritten in place has to be rejected too.
    const centralStart = archive.readUInt32LE(archive.length - 22 + 16);
    const brokenSignaturePath = path.join(dir, 'broken-signature.zip');
    const brokenSignature = Buffer.from(archive);
    brokenSignature.writeUInt32LE(0xdeadbeef, centralStart);
    fs.writeFileSync(brokenSignaturePath, brokenSignature);
    assert.equal(readDictionaryZipRevision(brokenSignaturePath), null);

    const brokenLengthPath = path.join(dir, 'broken-length.zip');
    const brokenLength = Buffer.from(archive);
    // Name length that runs the walk past the end of the directory.
    brokenLength.writeUInt16LE(0xffff, centralStart + 28);
    fs.writeFileSync(brokenLengthPath, brokenLength);
    assert.equal(readDictionaryZipRevision(brokenLengthPath), null);

    const foreignPath = path.join(dir, 'foreign.zip');
    fs.writeFileSync(foreignPath, Buffer.from('not a zip at all', 'utf8'));
    assert.equal(readDictionaryZipRevision(foreignPath), null);
  } finally {
    cleanupDir(dir);
  }
});
