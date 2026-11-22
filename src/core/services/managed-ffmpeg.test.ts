import assert from 'node:assert/strict';
import { chmod, mkdir, mkdtemp, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  appendManagedFfmpegToPath,
  detectFfmpeg,
  getManagedFfmpegDir,
  installManagedFfmpeg,
} from './managed-ffmpeg';

async function withTempDirs(run: (dirs: { system: string; managed: string }) => Promise<void>) {
  const root = await mkdtemp(path.join(tmpdir(), 'managed-ffmpeg-'));
  const dirs = { system: path.join(root, 'bin'), managed: getManagedFfmpegDir(root) };
  await mkdir(dirs.system, { recursive: true });
  await mkdir(dirs.managed, { recursive: true });
  try {
    await run(dirs);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
}

async function writeExecutable(directory: string, name: string): Promise<void> {
  const filePath = path.join(directory, name);
  await writeFile(filePath, '#!/bin/sh\n');
  await chmod(filePath, 0o755);
}

const passVersionCheck = async () => undefined;

test('managed FFmpeg is appended to PATH once, behind system entries, only when downloaded', async () => {
  await withTempDirs(async ({ system, managed }) => {
    const env: NodeJS.ProcessEnv = { Path: system };
    assert.equal(appendManagedFfmpegToPath(env, managed, 'linux'), false);

    await writeExecutable(managed, 'ffmpeg');
    assert.equal(appendManagedFfmpegToPath(env, managed, 'linux'), true);
    assert.equal(appendManagedFfmpegToPath(env, managed, 'linux'), false);
    // Writes the existing key (Windows spells it Path) instead of adding a second one.
    assert.deepEqual(env, { Path: [system, managed].join(path.delimiter) });
  });
});

test('detectFfmpeg prefers system tools and fills gaps from the managed copy', async () => {
  await withTempDirs(async ({ system, managed }) => {
    const env = { PATH: [system, managed].join(path.delimiter) };
    const detect = () =>
      detectFfmpeg({
        managedDir: managed,
        env,
        platform: 'linux',
        arch: 'x64',
        runVersionCheck: passVersionCheck,
      });

    assert.deepEqual(await detect(), {
      kind: 'missing',
      missing: ['ffmpeg', 'ffprobe'],
      downloadable: true,
    });

    await writeExecutable(system, 'ffmpeg');
    await writeExecutable(managed, 'ffprobe');
    assert.deepEqual(await detect(), {
      kind: 'ready',
      source: 'managed',
      ffmpegPath: path.join(system, 'ffmpeg'),
      ffprobePath: path.join(managed, 'ffprobe'),
    });

    await writeExecutable(system, 'ffprobe');
    const status = await detect();
    assert.equal(status.kind === 'ready' ? status.source : status.kind, 'system');
  });
});

test('detectFfmpeg reports a tool that is found but fails to run', async () => {
  await withTempDirs(async ({ managed }) => {
    await writeExecutable(managed, 'ffmpeg');
    await writeExecutable(managed, 'ffprobe');
    const status = await detectFfmpeg({
      managedDir: managed,
      env: { PATH: managed },
      platform: 'freebsd',
      arch: 'x64',
      runVersionCheck: async () => {
        throw new Error('Killed: 9');
      },
    });
    assert.deepEqual(status, {
      kind: 'broken',
      path: path.join(managed, 'ffmpeg'),
      message: 'Killed: 9',
      managed: true,
      downloadable: false,
    });
  });
});

test('installManagedFfmpeg rejects unsupported platforms and never publishes a failed download', async () => {
  await withTempDirs(async ({ managed }) => {
    await assert.rejects(
      installManagedFfmpeg({ managedDir: managed, platform: 'freebsd', arch: 'x64' }),
      /no FFmpeg download for freebsd-x64/,
    );

    const originalFetch = globalThis.fetch;
    globalThis.fetch = Object.assign(async () => new Response('not ffmpeg'), originalFetch);
    try {
      await assert.rejects(
        installManagedFfmpeg({ managedDir: managed, platform: 'linux', arch: 'x64' }),
        /FFmpeg failed integrity verification/,
      );
    } finally {
      globalThis.fetch = originalFetch;
    }
    assert.deepEqual(await readdir(managed), []);
  });
});
