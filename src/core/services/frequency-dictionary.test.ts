import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

import { createFrequencyDictionaryLookup } from './frequency-dictionary';

test('createFrequencyDictionaryLookup logs parse errors and returns no-op for invalid dictionaries', async () => {
  const logs: string[] = [];
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-frequency-dict-'));
  const bankPath = path.join(tempDir, 'term_meta_bank_1.json');
  fs.writeFileSync(bankPath, '{ invalid json');

  const lookup = await createFrequencyDictionaryLookup({
    searchPaths: [tempDir],
    log: (message) => {
      logs.push(message);
    },
  });

  const rank = lookup('猫');

  assert.equal(rank, null);
  assert.equal(
    logs.some(
      (entry) =>
        entry.includes('Failed to parse frequency dictionary file as JSON') &&
        entry.includes('term_meta_bank_1.json'),
    ),
    true,
  );
});

test('createFrequencyDictionaryLookup continues with no-op lookup when search path is missing', async () => {
  const logs: string[] = [];
  const missingPath = path.join(os.tmpdir(), 'subminer-frequency-dict-missing-dir');
  const lookup = await createFrequencyDictionaryLookup({
    searchPaths: [missingPath],
    log: (message) => {
      logs.push(message);
    },
  });

  assert.equal(lookup('猫'), null);
  assert.equal(
    logs.some((entry) => entry.includes(`Frequency dictionary not found.`)),
    true,
  );
});

test('createFrequencyDictionaryLookup aggregates duplicate-term logs into a single summary', async () => {
  const logs: string[] = [];
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-frequency-dict-'));
  const bankPath = path.join(tempDir, 'term_meta_bank_1.json');
  fs.writeFileSync(
    bankPath,
    JSON.stringify([
      ['猫', 1, { frequency: { displayValue: 100 } }],
      ['猫', 2, { frequency: { displayValue: 120 } }],
      ['猫', 3, { frequency: { displayValue: 110 } }],
    ]),
  );

  const lookup = await createFrequencyDictionaryLookup({
    searchPaths: [tempDir],
    log: (message) => {
      logs.push(message);
    },
  });

  assert.equal(lookup('猫'), 100);
  assert.equal(
    logs.filter((entry) => entry.includes('Frequency dictionary ignored 2 duplicate term entries'))
      .length,
    1,
  );
  assert.equal(
    logs.some((entry) => entry.includes('Frequency dictionary duplicate term')),
    false,
  );
});

test('createFrequencyDictionaryLookup prefers frequency.displayValue over value when both exist', async () => {
  const logs: string[] = [];
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-frequency-dict-'));
  const bankPath = path.join(tempDir, 'term_meta_bank_1.json');
  fs.writeFileSync(
    bankPath,
    JSON.stringify([
      ['猫', 1, { frequency: { value: 1234, displayValue: 1200 } }],
      ['鍛える', 2, { frequency: { value: 46961, displayValue: 2847 } }],
      ['犬', 2, { frequency: { displayValue: 88 } }],
    ]),
  );

  const lookup = await createFrequencyDictionaryLookup({
    searchPaths: [tempDir],
    log: (message) => {
      logs.push(message);
    },
  });

  assert.equal(lookup('猫'), 1200);
  assert.equal(lookup('鍛える'), 2847);
  assert.equal(lookup('犬'), 88);
  assert.equal(
    logs.some((entry) => entry.includes('Frequency dictionary loaded from')),
    true,
  );
});

test('createFrequencyDictionaryLookup parses composite displayValue by primary rank', async () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-frequency-dict-'));
  const bankPath = path.join(tempDir, 'term_meta_bank_1.json');
  fs.writeFileSync(
    bankPath,
    JSON.stringify([
      ['鍛える', 1, { frequency: { displayValue: '3272,52377' } }],
      ['高み', 2, { frequency: { displayValue: '9933,108961' } }],
    ]),
  );

  const lookup = await createFrequencyDictionaryLookup({
    searchPaths: [tempDir],
    log: () => undefined,
  });

  assert.equal(lookup('鍛える'), 3272);
  assert.equal(lookup('高み'), 9933);
});

test('createFrequencyDictionaryLookup does not require synchronous fs APIs', async () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-frequency-dict-'));
  const bankPath = path.join(tempDir, 'term_meta_bank_1.json');
  fs.writeFileSync(bankPath, JSON.stringify([['猫', 1, { frequency: { displayValue: 42 } }]]));

  const readFileSync = fs.readFileSync;
  const readdirSync = fs.readdirSync;
  const statSync = fs.statSync;
  const existsSync = fs.existsSync;
  (fs as unknown as Record<string, unknown>).readFileSync = () => {
    throw new Error('sync read disabled');
  };
  (fs as unknown as Record<string, unknown>).readdirSync = () => {
    throw new Error('sync readdir disabled');
  };
  (fs as unknown as Record<string, unknown>).statSync = () => {
    throw new Error('sync stat disabled');
  };
  (fs as unknown as Record<string, unknown>).existsSync = () => {
    throw new Error('sync exists disabled');
  };

  try {
    const lookup = await createFrequencyDictionaryLookup({
      searchPaths: [tempDir],
      log: () => undefined,
    });
    assert.equal(lookup('猫'), 42);
  } finally {
    (fs as unknown as Record<string, unknown>).readFileSync = readFileSync;
    (fs as unknown as Record<string, unknown>).readdirSync = readdirSync;
    (fs as unknown as Record<string, unknown>).statSync = statSync;
    (fs as unknown as Record<string, unknown>).existsSync = existsSync;
  }
});
