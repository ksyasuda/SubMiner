import assert from 'node:assert/strict';
import test from 'node:test';
import { EventEmitter } from 'node:events';
import * as childProcess from 'node:child_process';
import { PassThrough, Writable } from 'node:stream';
import { MecabTokenizer } from './mecab-tokenizer';

function createFakeMecabProcess(onKill: () => void): ReturnType<typeof childProcess.spawn> {
  const stdout = new PassThrough();
  const stderr = new PassThrough();
  const stdin = new Writable({
    write(chunk, _encoding, callback) {
      const text = String(chunk).replace(/\n+$/, '').trim();
      if (!text) {
        stdout.write('EOS\n');
        callback();
        return;
      }

      const payload = `${text}\t名詞,一般,*,*,*,*,${text},${text},${text}\nEOS\n`;
      stdout.write(payload);
      callback();
    },
  });

  const process = new EventEmitter() as unknown as ReturnType<typeof childProcess.spawn> & {
    stdin: Writable;
    stdout: PassThrough;
    stderr: PassThrough;
  };
  process.stdin = stdin;
  process.stdout = stdout;
  process.stderr = stderr;
  process.kill = () => {
    onKill();
    process.emit('close', 0);
    return true;
  };
  return process;
}

test('MecabTokenizer reuses a persistent parser process across subtitle lines', async () => {
  let spawnCalls = 0;
  let killCalls = 0;
  let timerId = 0;
  const timers = new Map<number, () => void>();

  const tokenizer = new MecabTokenizer({
    execSyncFn: (() => '/usr/bin/mecab') as unknown as typeof childProcess.execSync,
    spawnFn: (() => {
      spawnCalls += 1;
      return createFakeMecabProcess(() => {
        killCalls += 1;
      });
    }) as unknown as typeof childProcess.spawn,
    setTimeoutFn: (callback) => {
      timerId += 1;
      timers.set(timerId, callback);
      return timerId as unknown as ReturnType<typeof setTimeout>;
    },
    clearTimeoutFn: (timeout) => {
      timers.delete(timeout as unknown as number);
    },
    idleShutdownMs: 60_000,
  });

  assert.equal(await tokenizer.checkAvailability(), true);

  const first = await tokenizer.tokenize('猫');
  const second = await tokenizer.tokenize('犬');

  assert.equal(first?.[0]?.word, '猫');
  assert.equal(second?.[0]?.word, '犬');
  assert.equal(spawnCalls, 1);
  assert.equal(killCalls, 0);
});

test('MecabTokenizer shuts down after idle timeout and restarts on new activity', async () => {
  let spawnCalls = 0;
  let killCalls = 0;
  let timerId = 0;
  const timers = new Map<number, () => void>();

  const tokenizer = new MecabTokenizer({
    execSyncFn: (() => '/usr/bin/mecab') as unknown as typeof childProcess.execSync,
    spawnFn: (() => {
      spawnCalls += 1;
      return createFakeMecabProcess(() => {
        killCalls += 1;
      });
    }) as unknown as typeof childProcess.spawn,
    setTimeoutFn: (callback) => {
      timerId += 1;
      timers.set(timerId, callback);
      return timerId as unknown as ReturnType<typeof setTimeout>;
    },
    clearTimeoutFn: (timeout) => {
      timers.delete(timeout as unknown as number);
    },
    idleShutdownMs: 5_000,
  });

  assert.equal(await tokenizer.checkAvailability(), true);
  await tokenizer.tokenize('猫');
  assert.equal(spawnCalls, 1);

  const pendingTimer = [...timers.values()][0];
  assert.ok(pendingTimer, 'expected idle shutdown timer');
  pendingTimer?.();
  assert.equal(killCalls, 1);

  await tokenizer.tokenize('犬');
  assert.equal(spawnCalls, 2);
});
