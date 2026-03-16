import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { createJlptVocabularyLookup } from './jlpt-vocab';

test('createJlptVocabularyLookup loads JLPT bank entries and resolves known levels', async () => {
  const logs: string[] = [];
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-jlpt-dict-'));
  fs.writeFileSync(
    path.join(tempDir, 'term_meta_bank_5.json'),
    JSON.stringify([
      ['猫', 1, { frequency: { displayValue: 1 } }],
      ['犬', 2, { frequency: { displayValue: 2 } }],
    ]),
  );
  fs.writeFileSync(path.join(tempDir, 'term_meta_bank_1.json'), JSON.stringify([]));
  fs.writeFileSync(path.join(tempDir, 'term_meta_bank_2.json'), JSON.stringify([]));
  fs.writeFileSync(path.join(tempDir, 'term_meta_bank_3.json'), JSON.stringify([]));
  fs.writeFileSync(path.join(tempDir, 'term_meta_bank_4.json'), JSON.stringify([]));

  const lookup = await createJlptVocabularyLookup({
    searchPaths: [tempDir],
    log: (message) => {
      logs.push(message);
    },
  });

  assert.equal(lookup('猫'), 'N5');
  assert.equal(lookup('犬'), 'N5');
  assert.equal(lookup('鳥'), null);
  assert.equal(
    logs.some((entry) => entry.includes('JLPT dictionary loaded from')),
    true,
  );
});

test('createJlptVocabularyLookup does not require synchronous fs APIs', async () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-jlpt-dict-'));
  fs.writeFileSync(
    path.join(tempDir, 'term_meta_bank_4.json'),
    JSON.stringify([['見る', 1, { frequency: { displayValue: 3 } }]]),
  );
  fs.writeFileSync(path.join(tempDir, 'term_meta_bank_1.json'), JSON.stringify([]));
  fs.writeFileSync(path.join(tempDir, 'term_meta_bank_2.json'), JSON.stringify([]));
  fs.writeFileSync(path.join(tempDir, 'term_meta_bank_3.json'), JSON.stringify([]));
  fs.writeFileSync(path.join(tempDir, 'term_meta_bank_5.json'), JSON.stringify([]));

  const readFileSync = fs.readFileSync;
  const statSync = fs.statSync;
  const existsSync = fs.existsSync;
  (fs as unknown as Record<string, unknown>).readFileSync = () => {
    throw new Error('sync read disabled');
  };
  (fs as unknown as Record<string, unknown>).statSync = () => {
    throw new Error('sync stat disabled');
  };
  (fs as unknown as Record<string, unknown>).existsSync = () => {
    throw new Error('sync exists disabled');
  };

  try {
    const lookup = await createJlptVocabularyLookup({
      searchPaths: [tempDir],
      log: () => undefined,
    });
    assert.equal(lookup('見る'), 'N4');
  } finally {
    (fs as unknown as Record<string, unknown>).readFileSync = readFileSync;
    (fs as unknown as Record<string, unknown>).statSync = statSync;
    (fs as unknown as Record<string, unknown>).existsSync = existsSync;
  }
});

test('createJlptVocabularyLookup summarizes duplicate JLPT terms without per-entry log spam', async () => {
  const logs: string[] = [];
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-jlpt-dict-'));
  fs.writeFileSync(
    path.join(tempDir, 'term_meta_bank_1.json'),
    JSON.stringify([
      ['余り', 1, { frequency: { displayValue: 'N1' }, reading: 'あんまり' }],
      ['私', 2, { frequency: { displayValue: 'N1' }, reading: 'あたし' }],
    ]),
  );
  fs.writeFileSync(path.join(tempDir, 'term_meta_bank_2.json'), JSON.stringify([]));
  fs.writeFileSync(path.join(tempDir, 'term_meta_bank_3.json'), JSON.stringify([]));
  fs.writeFileSync(path.join(tempDir, 'term_meta_bank_4.json'), JSON.stringify([]));
  fs.writeFileSync(
    path.join(tempDir, 'term_meta_bank_5.json'),
    JSON.stringify([
      ['余り', 3, { frequency: { displayValue: 'N5' }, reading: 'あまり' }],
      ['私', 4, { frequency: { displayValue: 'N5' }, reading: 'わたし' }],
      ['私', 5, { frequency: { displayValue: 'N5' }, reading: 'わたくし' }],
    ]),
  );

  const lookup = await createJlptVocabularyLookup({
    searchPaths: [tempDir],
    log: (message) => {
      logs.push(message);
    },
  });

  assert.equal(lookup('余り'), 'N1');
  assert.equal(lookup('私'), 'N1');
  assert.equal(
    logs.some((entry) => entry.includes('keeping that level instead of')),
    false,
  );
  assert.equal(
    logs.some((entry) => entry.includes('collapsed 3 duplicate JLPT entries across 2 terms')),
    true,
  );
});
