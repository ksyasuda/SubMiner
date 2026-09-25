import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createTransferCache, transferCacheKey } from './transfer-cache';

test('transfer cache isolates peers and active transfers while replacing previous snapshots', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-cache-test-'));
  try {
    const cacheDir = path.join(root, 'cache');
    const cache = createTransferCache(cacheDir);
    const key = transferCacheKey('peer');
    const first = path.join(root, 'first');
    fs.mkdirSync(path.join(first, 'incoming'), { recursive: true });
    const incoming = path.join(first, 'incoming', 'snapshot.sqlite');
    fs.writeFileSync(incoming, 'first received snapshot');
    cache.remember(key, first);
    const second = path.join(root, 'second');
    cache.seed(key, second);
    fs.writeFileSync(incoming, 'next received snapshot');
    cache.remember(key, first);
    assert.equal(
      fs.readFileSync(path.join(second, 'incoming', 'snapshot.sqlite'), 'utf8'),
      'first received snapshot',
    );
    const third = path.join(root, 'third');
    cache.seed(key, third);
    assert.equal(
      fs.readFileSync(path.join(third, 'incoming', 'snapshot.sqlite'), 'utf8'),
      'next received snapshot',
    );
    assert.deepEqual(fs.readdirSync(cacheDir), [`${key}.sqlite`]);
    const other = path.join(root, 'other');
    cache.seed(transferCacheKey('other peer'), other);
    assert.equal(fs.existsSync(path.join(other, 'incoming', 'snapshot.sqlite')), false);
    assert.throws(() => cache.seed('../outside', other), /Invalid/);
    assert.throws(() => cache.remember('../outside', first), /Invalid/);
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test('unavailable cache storage and missing incoming snapshots do not prevent sync', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-cache-test-'));
  try {
    const unavailable = path.join(root, 'file');
    fs.writeFileSync(unavailable, 'not a directory');
    const cache = createTransferCache(unavailable);
    const key = transferCacheKey('peer');
    const temp = path.join(root, 'transfer');
    assert.doesNotThrow(() => cache.seed(key, temp));
    assert.doesNotThrow(() => cache.remember(key, temp));
    fs.writeFileSync(path.join(temp, 'incoming', 'snapshot.sqlite'), 'received');
    assert.doesNotThrow(() => cache.remember(key, temp));
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});
