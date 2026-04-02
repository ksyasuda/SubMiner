import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildWindowsMpvLaunchArgs,
  launchWindowsMpv,
  resolveWindowsMpvPath,
  type WindowsMpvLaunchDeps,
} from './windows-mpv-launch';

function createDeps(overrides: Partial<WindowsMpvLaunchDeps> = {}): WindowsMpvLaunchDeps {
  return {
    getEnv: () => undefined,
    runWhere: () => ({ status: 1, stdout: '' }),
    fileExists: () => false,
    spawnDetached: () => undefined,
    showError: () => undefined,
    ...overrides,
  };
}

test('resolveWindowsMpvPath prefers SUBMINER_MPV_PATH', () => {
  const resolved = resolveWindowsMpvPath(
    createDeps({
      getEnv: (name) => (name === 'SUBMINER_MPV_PATH' ? 'C:\\mpv\\mpv.exe' : undefined),
      fileExists: (candidate) => candidate === 'C:\\mpv\\mpv.exe',
    }),
  );

  assert.equal(resolved, 'C:\\mpv\\mpv.exe');
});

test('resolveWindowsMpvPath falls back to where.exe output', () => {
  const resolved = resolveWindowsMpvPath(
    createDeps({
      runWhere: () => ({ status: 0, stdout: 'C:\\tools\\mpv.exe\r\nC:\\other\\mpv.exe\r\n' }),
      fileExists: (candidate) => candidate === 'C:\\tools\\mpv.exe',
    }),
  );

  assert.equal(resolved, 'C:\\tools\\mpv.exe');
});

test('buildWindowsMpvLaunchArgs uses explicit SubMiner defaults and targets', () => {
  assert.deepEqual(buildWindowsMpvLaunchArgs(['C:\\a.mkv', 'C:\\b.mkv']), [
    '--player-operation-mode=pseudo-gui',
    '--input-ipc-server=\\\\.\\pipe\\subminer-socket',
    '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
    '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
    '--sub-auto=fuzzy',
    '--sub-file-paths=.;subs;subtitles',
    '--sid=auto',
    '--secondary-sid=auto',
    '--secondary-sub-visibility=no',
    'C:\\a.mkv',
    'C:\\b.mkv',
  ]);
});

test('launchWindowsMpv reports missing mpv path', () => {
  const errors: string[] = [];
  const result = launchWindowsMpv(
    [],
    createDeps({
      showError: (_title, content) => errors.push(content),
    }),
  );

  assert.equal(result.ok, false);
  assert.equal(result.mpvPath, '');
  assert.match(errors[0] ?? '', /Could not find mpv\.exe/i);
});

test('launchWindowsMpv spawns detached mpv with targets', () => {
  const calls: string[] = [];
  const result = launchWindowsMpv(
    ['C:\\video.mkv'],
    createDeps({
      getEnv: (name) => (name === 'SUBMINER_MPV_PATH' ? 'C:\\mpv\\mpv.exe' : undefined),
      fileExists: (candidate) => candidate === 'C:\\mpv\\mpv.exe',
      spawnDetached: (command, args) => {
        calls.push(command);
        calls.push(args.join('|'));
      },
    }),
  );

  assert.equal(result.ok, true);
  assert.equal(result.mpvPath, 'C:\\mpv\\mpv.exe');
  assert.deepEqual(calls, [
    'C:\\mpv\\mpv.exe',
    '--player-operation-mode=pseudo-gui|--input-ipc-server=\\\\.\\pipe\\subminer-socket|--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us|--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us|--sub-auto=fuzzy|--sub-file-paths=.;subs;subtitles|--sid=auto|--secondary-sid=auto|--secondary-sub-visibility=no|C:\\video.mkv',
  ]);
});

test('launchWindowsMpv reports spawn failures with path context', () => {
  const errors: string[] = [];
  const result = launchWindowsMpv(
    [],
    createDeps({
      getEnv: (name) => (name === 'SUBMINER_MPV_PATH' ? 'C:\\mpv\\mpv.exe' : undefined),
      fileExists: (candidate) => candidate === 'C:\\mpv\\mpv.exe',
      spawnDetached: () => {
        throw new Error('spawn failed');
      },
      showError: (_title, content) => errors.push(content),
    }),
  );

  assert.equal(result.ok, false);
  assert.equal(result.mpvPath, 'C:\\mpv\\mpv.exe');
  assert.match(errors[0] ?? '', /Failed to launch mpv/i);
  assert.match(errors[0] ?? '', /C:\\mpv\\mpv\.exe/i);
});
