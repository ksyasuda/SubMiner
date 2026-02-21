import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { resolveConfigFilePath } from '../src/config/path-resolution.js';

type RunResult = {
  status: number | null;
  stdout: string;
  stderr: string;
};

function withTempDir<T>(fn: (dir: string) => T): T {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-launcher-test-'));
  try {
    return fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function runLauncher(argv: string[], env: NodeJS.ProcessEnv): RunResult {
  const result = spawnSync(process.execPath, ['run', path.join(process.cwd(), 'launcher/main.ts'), ...argv], {
    env,
    encoding: 'utf8',
  });
  return {
    status: result.status,
    stdout: result.stdout || '',
    stderr: result.stderr || '',
  };
}

function makeTestEnv(homeDir: string, xdgConfigHome: string): NodeJS.ProcessEnv {
  return {
    ...process.env,
    HOME: homeDir,
    XDG_CONFIG_HOME: xdgConfigHome,
  };
}

test('config path uses XDG_CONFIG_HOME override', () => {
  withTempDir((root) => {
    const xdgConfigHome = path.join(root, 'xdg');
    const homeDir = path.join(root, 'home');
    fs.mkdirSync(path.join(xdgConfigHome, 'SubMiner'), { recursive: true });
    fs.writeFileSync(path.join(xdgConfigHome, 'SubMiner', 'config.json'), '{"source":"xdg"}');

    const result = runLauncher(['config', 'path'], makeTestEnv(homeDir, xdgConfigHome));

    assert.equal(result.status, 0);
    assert.equal(result.stdout.trim(), path.join(xdgConfigHome, 'SubMiner', 'config.json'));
  });
});

test('config discovery ignores lowercase subminer candidate', () => {
  const homeDir = '/home/tester';
  const xdgConfigHome = '/tmp/xdg-config';
  const expected = path.join(xdgConfigHome, 'SubMiner', 'config.jsonc');
  const foundPaths = new Set([path.join(xdgConfigHome, 'subminer', 'config.json')]);

  const resolved = resolveConfigFilePath({
    xdgConfigHome,
    homeDir,
    existsSync: (candidate) => foundPaths.has(path.normalize(candidate)),
  });

  assert.equal(resolved, expected);
});

test('config path prefers jsonc over json for same directory', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    fs.mkdirSync(path.join(xdgConfigHome, 'SubMiner'), { recursive: true });
    fs.writeFileSync(path.join(xdgConfigHome, 'SubMiner', 'config.json'), '{"format":"json"}');
    fs.writeFileSync(path.join(xdgConfigHome, 'SubMiner', 'config.jsonc'), '{"format":"jsonc"}');

    const result = runLauncher(['config', 'path'], makeTestEnv(homeDir, xdgConfigHome));

    assert.equal(result.status, 0);
    assert.equal(result.stdout.trim(), path.join(xdgConfigHome, 'SubMiner', 'config.jsonc'));
  });
});

test('config show prints config body and appends trailing newline', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    fs.mkdirSync(path.join(xdgConfigHome, 'SubMiner'), { recursive: true });
    fs.writeFileSync(path.join(xdgConfigHome, 'SubMiner', 'config.jsonc'), '{"logLevel":"debug"}');

    const result = runLauncher(['config', 'show'], makeTestEnv(homeDir, xdgConfigHome));

    assert.equal(result.status, 0);
    assert.equal(result.stdout, '{"logLevel":"debug"}\n');
  });
});

test('mpv socket command returns socket path from plugin runtime config', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const expectedSocket = path.join(root, 'custom', 'subminer.sock');
    fs.mkdirSync(path.join(xdgConfigHome, 'mpv', 'script-opts'), { recursive: true });
    fs.writeFileSync(
      path.join(xdgConfigHome, 'mpv', 'script-opts', 'subminer.conf'),
      `socket_path=${expectedSocket}\n`,
    );

    const result = runLauncher(['mpv', 'socket'], makeTestEnv(homeDir, xdgConfigHome));

    assert.equal(result.status, 0);
    assert.equal(result.stdout.trim(), expectedSocket);
  });
});

test('mpv status exits non-zero when socket is not ready', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const result = runLauncher(['mpv', 'status'], makeTestEnv(homeDir, xdgConfigHome));

    assert.equal(result.status, 1);
    assert.match(result.stdout, /socket not ready/i);
  });
});

test('doctor reports checks and exits non-zero without hard dependencies', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const env = {
      ...makeTestEnv(homeDir, xdgConfigHome),
      PATH: '',
    };
    const result = runLauncher(['doctor'], env);

    assert.equal(result.status, 1);
    assert.match(result.stdout, /\[doctor\] app binary:/);
    assert.match(result.stdout, /\[doctor\] mpv:/);
    assert.match(result.stdout, /\[doctor\] config:/);
  });
});

test('jellyfin discovery routes to app --start with log-level forwarding', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const appPath = path.join(root, 'fake-subminer.sh');
    const capturePath = path.join(root, 'captured-args.txt');
    fs.writeFileSync(
      appPath,
      '#!/bin/sh\nif [ -n "$SUBMINER_TEST_CAPTURE" ]; then printf "%s\\n" "$@" > "$SUBMINER_TEST_CAPTURE"; fi\nexit 0\n',
    );
    fs.chmodSync(appPath, 0o755);

    const env = {
      ...makeTestEnv(homeDir, xdgConfigHome),
      SUBMINER_APPIMAGE_PATH: appPath,
      SUBMINER_TEST_CAPTURE: capturePath,
    };
    const result = runLauncher(['jellyfin', 'discovery', '--log-level', 'debug'], env);

    assert.equal(result.status, 0);
    assert.equal(fs.readFileSync(capturePath, 'utf8'), '--start\n--log-level\ndebug\n');
  });
});

test('jellyfin login routes credentials to app command', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const appPath = path.join(root, 'fake-subminer.sh');
    const capturePath = path.join(root, 'captured-args.txt');
    fs.writeFileSync(
      appPath,
      '#!/bin/sh\nif [ -n "$SUBMINER_TEST_CAPTURE" ]; then printf "%s\\n" "$@" > "$SUBMINER_TEST_CAPTURE"; fi\nexit 0\n',
    );
    fs.chmodSync(appPath, 0o755);

    const env = {
      ...makeTestEnv(homeDir, xdgConfigHome),
      SUBMINER_APPIMAGE_PATH: appPath,
      SUBMINER_TEST_CAPTURE: capturePath,
    };
    const result = runLauncher(
      [
        'jellyfin',
        'login',
        '--server',
        'https://jf.example.test',
        '--username',
        'alice',
        '--password',
        'secret',
      ],
      env,
    );

    assert.equal(result.status, 0);
    assert.equal(
      fs.readFileSync(capturePath, 'utf8'),
      '--jellyfin-login\n--jellyfin-server\nhttps://jf.example.test\n--jellyfin-username\nalice\n--jellyfin-password\nsecret\n',
    );
  });
});
