import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { resolveConfigFilePath } from '../src/config/path-resolution.js';
import {
  parseJellyfinLibrariesFromAppOutput,
  parseJellyfinItemsFromAppOutput,
  parseJellyfinErrorFromAppOutput,
  parseJellyfinPreviewAuthResponse,
  deriveJellyfinTokenStorePath,
  hasStoredJellyfinSession,
  shouldRetryWithStartForNoRunningInstance,
  readUtf8FileAppendedSince,
  parseEpisodePathFromDisplay,
  buildRootSearchGroups,
  classifyJellyfinChildSelection,
  buildForwardedJellyfinAppArgs,
} from './jellyfin.js';

type RunResult = {
  status: number | null;
  stdout: string;
  stderr: string;
};

function withTempDir<T>(fn: (dir: string) => T): T {
  // Keep paths short on macOS/Linux: Unix domain sockets have small path-length limits.
  const tmpBase = process.platform === 'win32' ? os.tmpdir() : '/tmp';
  const dir = fs.mkdtempSync(path.join(tmpBase, 'subminer-launcher-test-'));
  try {
    return fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

const LAUNCHER_RUN_TIMEOUT_MS = 30000;

function runLauncher(argv: string[], env: NodeJS.ProcessEnv): RunResult {
  const result = spawnSync(
    process.execPath,
    ['run', path.join(process.cwd(), 'launcher/main.ts'), ...argv],
    {
      env,
      encoding: 'utf8',
      timeout: LAUNCHER_RUN_TIMEOUT_MS,
    },
  );
  return {
    status: result.status,
    stdout: result.stdout || '',
    stderr: result.stderr || '',
  };
}

function makeTestEnv(homeDir: string, xdgConfigHome: string): NodeJS.ProcessEnv {
  const pathValue = process.env.Path || process.env.PATH || '';
  return {
    ...process.env,
    HOME: homeDir,
    USERPROFILE: homeDir,
    APPDATA: xdgConfigHome,
    LOCALAPPDATA: path.join(homeDir, 'AppData', 'Local'),
    XDG_CONFIG_HOME: xdgConfigHome,
    PATH: pathValue,
    Path: pathValue,
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
  const expected = path.posix.join(xdgConfigHome, 'SubMiner', 'config.jsonc');
  const foundPaths = new Set([path.posix.join(xdgConfigHome, 'subminer', 'config.json')]);

  const resolved = resolveConfigFilePath({
    xdgConfigHome,
    homeDir,
    platform: 'linux',
    existsSync: (candidate) => foundPaths.has(path.posix.normalize(candidate)),
  });

  assert.equal(resolved, expected);
});

test('version flag prints installed app version without requiring app binary', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const result = runLauncher(['--version'], makeTestEnv(homeDir, xdgConfigHome));

    assert.equal(result.status, 0);
    assert.match(result.stdout.trim(), /^SubMiner \d+\.\d+\.\d+/);
    assert.equal(result.stderr, '');
  });
});

test('short version flag prints installed app version without requiring app binary', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const result = runLauncher(['-v'], makeTestEnv(homeDir, xdgConfigHome));

    assert.equal(result.status, 0);
    assert.match(result.stdout.trim(), /^SubMiner \d+\.\d+\.\d+/);
    assert.equal(result.stderr, '');
  });
});

test('logs export writes sanitized archive without requiring app binary', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const logsDir =
      process.platform === 'win32'
        ? path.join(xdgConfigHome, 'SubMiner', 'logs')
        : path.join(homeDir, '.config', 'SubMiner', 'logs');
    fs.mkdirSync(logsDir, { recursive: true });
    fs.writeFileSync(path.join(logsDir, 'app-2026-W21.log'), `/home/kyle/video.mkv\n`, 'utf8');

    const result = runLauncher(['logs', '-e'], makeTestEnv(homeDir, xdgConfigHome));

    assert.equal(result.status, 0, `stdout:\n${result.stdout}\nstderr:\n${result.stderr}`);
    const zipPath = result.stdout.trim();
    assert.match(zipPath, /subminer-logs-.+\.zip$/);
    assert.equal(fs.existsSync(zipPath), true);
    const archive = fs.readFileSync(zipPath);
    assert.equal(archive.includes(Buffer.from('/home/kyle')), false);
    assert.equal(archive.includes(Buffer.from('/home/<user>')), true);
  });
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
    fs.mkdirSync(path.join(xdgConfigHome, 'SubMiner'), { recursive: true });
    fs.writeFileSync(
      path.join(xdgConfigHome, 'SubMiner', 'config.jsonc'),
      JSON.stringify({ mpv: { socketPath: expectedSocket } }),
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
    const socketPath = path.join(root, 'missing.sock');
    fs.mkdirSync(path.join(xdgConfigHome, 'SubMiner'), { recursive: true });
    fs.writeFileSync(
      path.join(xdgConfigHome, 'SubMiner', 'config.jsonc'),
      JSON.stringify({ mpv: { socketPath } }),
    );
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
      Path: '',
    };
    const result = runLauncher(['doctor'], env);

    assert.equal(result.status, 1);
    assert.match(result.stdout, /\[doctor\] app binary:/);
    assert.match(result.stdout, /\[doctor\] mpv:/);
    assert.match(result.stdout, /\[doctor\] config:/);
  });
});

test('doctor refresh-known-words forwards app refresh command without requiring mpv', () => {
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
      PATH: '',
      Path: '',
      SUBMINER_APPIMAGE_PATH: appPath,
      SUBMINER_TEST_CAPTURE: capturePath,
    };
    const result = runLauncher(['doctor', '--refresh-known-words'], env);

    assert.equal(result.status, 0);
    assert.equal(fs.readFileSync(capturePath, 'utf8'), '--refresh-known-words\n');
    assert.match(result.stdout, /\[doctor\] mpv: missing/);
  });
});

test('launcher settings option forwards app settings window command', () => {
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
    const result = runLauncher(['--settings'], env);

    assert.equal(result.status, 0);
    assert.equal(fs.readFileSync(capturePath, 'utf8'), '--settings\n');
  });
});

test('launcher settings command forwards app settings window command', () => {
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
    const result = runLauncher(['settings'], env);

    assert.equal(result.status, 0);
    assert.equal(fs.readFileSync(capturePath, 'utf8'), '--settings\n');
  });
});

test('launcher settings command suppresses known Electron macOS menu diagnostics', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const appPath = path.join(root, 'fake-subminer.sh');
    fs.writeFileSync(
      appPath,
      [
        '#!/bin/sh',
        'printf "%s\\n" "2026-05-17 02:59:52.141 SubMiner[29060:305323] representedObject is not a WeakPtrToElectronMenuModelAsNSObject" >&2',
        'printf "%s\\n" "real stderr line" >&2',
        'exit 0',
        '',
      ].join('\n'),
    );
    fs.chmodSync(appPath, 0o755);

    const env = {
      ...makeTestEnv(homeDir, xdgConfigHome),
      SUBMINER_APPIMAGE_PATH: appPath,
    };
    const result = runLauncher(['settings'], env);

    assert.equal(result.status, 0);
    assert.equal(result.stderr, 'real stderr line\n');
  });
});

test('launcher forwards --args to mpv as parsed tokens', { timeout: 15000 }, () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const binDir = path.join(root, 'bin');
    const appPath = path.join(root, 'fake-subminer.sh');
    const videoPath = path.join(root, 'movie.mkv');
    const mpvArgsPath = path.join(root, 'mpv-args.txt');
    const socketPath = path.join(root, 'mpv.sock');
    const bunBinary = JSON.stringify(process.execPath.replace(/\\/g, '/'));

    fs.mkdirSync(binDir, { recursive: true });
    fs.mkdirSync(path.join(xdgConfigHome, 'SubMiner'), { recursive: true });
    fs.writeFileSync(videoPath, 'fake video content');
    fs.writeFileSync(
      path.join(xdgConfigHome, 'SubMiner', 'setup-state.json'),
      JSON.stringify({
        version: 1,
        status: 'completed',
        completedAt: '2026-03-08T00:00:00.000Z',
        completionSource: 'user',
        lastSeenYomitanDictionaryCount: 0,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
      }),
    );
    fs.writeFileSync(
      path.join(xdgConfigHome, 'SubMiner', 'config.jsonc'),
      JSON.stringify({
        auto_start_overlay: false,
        mpv: {
          socketPath,
          autoStartSubMiner: false,
          pauseUntilOverlayReady: false,
        },
      }),
    );
    fs.writeFileSync(appPath, '#!/bin/sh\nexit 0\n');
    fs.chmodSync(appPath, 0o755);

    fs.writeFileSync(
      path.join(binDir, 'mpv'),
      `#!/bin/sh
set -eu
printf '%s\\n' "$@" > "$SUBMINER_TEST_MPV_ARGS"
socket_path=""
for arg in "$@"; do
  case "$arg" in
    --input-ipc-server=*)
      socket_path="\${arg#--input-ipc-server=}"
      ;;
  esac
done
${bunBinary} -e "const net=require('node:net'); const fs=require('node:fs'); const path=require('node:path'); const socket=process.argv[1]||''; try{ if (socket) fs.mkdirSync(path.dirname(socket),{recursive:true}); }catch{} try{ if (socket) fs.rmSync(socket,{force:true}); }catch{} if(!socket) process.exit(0); const server=net.createServer((c)=>c.end()); server.on('error',()=>process.exit(0)); try{ server.listen(socket,()=>setTimeout(()=>server.close(()=>process.exit(0)),250)); } catch { process.exit(0); }" "$socket_path"
`,
      'utf8',
    );
    fs.chmodSync(path.join(binDir, 'mpv'), 0o755);
    fs.writeFileSync(path.join(binDir, 'yt-dlp'), '#!/bin/sh\nexit 0\n', 'utf8');
    fs.writeFileSync(path.join(binDir, 'ffmpeg'), '#!/bin/sh\nexit 0\n', 'utf8');
    fs.chmodSync(path.join(binDir, 'yt-dlp'), 0o755);
    fs.chmodSync(path.join(binDir, 'ffmpeg'), 0o755);

    const env = {
      ...makeTestEnv(homeDir, xdgConfigHome),
      PATH: `${binDir}${path.delimiter}${process.env.Path || process.env.PATH || ''}`,
      Path: `${binDir}${path.delimiter}${process.env.Path || process.env.PATH || ''}`,
      SUBMINER_APPIMAGE_PATH: appPath,
      SUBMINER_TEST_MPV_ARGS: mpvArgsPath,
    };
    const result = runLauncher(['--args', '--pause=yes --title="movie night"', videoPath], env);

    assert.equal(result.status, 0, `stdout:\n${result.stdout}\nstderr:\n${result.stderr}`);
    const argsFile = fs.readFileSync(mpvArgsPath, 'utf8');
    const forwardedArgs = argsFile
      .trim()
      .split('\n')
      .map((item) => item.trim())
      .filter(Boolean);

    assert.equal(forwardedArgs.includes('--pause=yes'), true);
    assert.equal(forwardedArgs.includes('--title=movie night'), true);
    assert.equal(forwardedArgs.includes(videoPath), true);
  });
});

test('launcher forwards non-info log level into mpv logging args', { timeout: 15000 }, () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const binDir = path.join(root, 'bin');
    const appPath = path.join(root, 'fake-subminer.sh');
    const videoPath = path.join(root, 'movie.mkv');
    const mpvArgsPath = path.join(root, 'mpv-args.txt');
    const socketPath = path.join(root, 'mpv.sock');
    const bunBinary = JSON.stringify(process.execPath.replace(/\\/g, '/'));

    fs.mkdirSync(binDir, { recursive: true });
    fs.mkdirSync(path.join(xdgConfigHome, 'SubMiner'), { recursive: true });
    fs.writeFileSync(videoPath, 'fake video content');
    fs.writeFileSync(
      path.join(xdgConfigHome, 'SubMiner', 'setup-state.json'),
      JSON.stringify({
        version: 1,
        status: 'completed',
        completedAt: '2026-03-08T00:00:00.000Z',
        completionSource: 'user',
        lastSeenYomitanDictionaryCount: 0,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
      }),
    );
    fs.writeFileSync(
      path.join(xdgConfigHome, 'SubMiner', 'config.jsonc'),
      JSON.stringify({
        auto_start_overlay: true,
        mpv: {
          socketPath,
          autoStartSubMiner: true,
          pauseUntilOverlayReady: true,
        },
        logging: {
          files: {
            mpv: true,
          },
        },
      }),
    );
    fs.writeFileSync(appPath, '#!/bin/sh\nexit 0\n');
    fs.chmodSync(appPath, 0o755);

    fs.writeFileSync(
      path.join(binDir, 'mpv'),
      `#!/bin/sh
set -eu
printf '%s\\n' "$@" > "$SUBMINER_TEST_MPV_ARGS"
socket_path=""
for arg in "$@"; do
  case "$arg" in
    --input-ipc-server=*)
      socket_path="\${arg#--input-ipc-server=}"
      ;;
  esac
done
${bunBinary} -e "const net=require('node:net'); const fs=require('node:fs'); const path=require('node:path'); const socket=process.argv[1]||''; try{ if (socket) fs.mkdirSync(path.dirname(socket),{recursive:true}); }catch{} try{ if (socket) fs.rmSync(socket,{force:true}); }catch{} if(!socket) process.exit(0); const server=net.createServer((c)=>c.end()); server.on('error',()=>process.exit(0)); try{ server.listen(socket,()=>setTimeout(()=>server.close(()=>process.exit(0)),250)); } catch { process.exit(0); }" "$socket_path"
`,
      'utf8',
    );
    fs.chmodSync(path.join(binDir, 'mpv'), 0o755);
    fs.writeFileSync(path.join(binDir, 'yt-dlp'), '#!/bin/sh\nexit 0\n', 'utf8');
    fs.writeFileSync(path.join(binDir, 'ffmpeg'), '#!/bin/sh\nexit 0\n', 'utf8');
    fs.chmodSync(path.join(binDir, 'yt-dlp'), 0o755);
    fs.chmodSync(path.join(binDir, 'ffmpeg'), 0o755);

    const env = {
      ...makeTestEnv(homeDir, xdgConfigHome),
      PATH: `${binDir}${path.delimiter}${process.env.Path || process.env.PATH || ''}`,
      Path: `${binDir}${path.delimiter}${process.env.Path || process.env.PATH || ''}`,
      SUBMINER_APPIMAGE_PATH: appPath,
      SUBMINER_TEST_MPV_ARGS: mpvArgsPath,
    };
    const result = runLauncher(['--log-level', 'debug', videoPath], env);

    assert.equal(result.status, 0, `stdout:\n${result.stdout}\nstderr:\n${result.stderr}`);
    const mpvArgs = fs.readFileSync(mpvArgsPath, 'utf8');
    assert.match(mpvArgs, /--msg-level=all=warn,subminer=debug/);
    assert.doesNotMatch(mpvArgs, /--script-opts=.*subminer-log_level=debug/);
  });
});

test('launcher routes youtube urls through regular playback startup', { timeout: 15000 }, () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const binDir = path.join(root, 'bin');
    const appPath = path.join(root, 'fake-subminer.sh');
    const mpvArgsPath = path.join(root, 'mpv-args.txt');
    const socketPath = path.join(root, 'mpv.sock');
    const bunBinary = JSON.stringify(process.execPath.replace(/\\/g, '/'));

    fs.mkdirSync(binDir, { recursive: true });
    fs.mkdirSync(path.join(xdgConfigHome, 'SubMiner'), { recursive: true });
    fs.writeFileSync(
      path.join(xdgConfigHome, 'SubMiner', 'setup-state.json'),
      JSON.stringify({
        version: 1,
        status: 'completed',
        completedAt: '2026-03-08T00:00:00.000Z',
        completionSource: 'user',
        lastSeenYomitanDictionaryCount: 0,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
      }),
    );
    fs.writeFileSync(
      path.join(xdgConfigHome, 'SubMiner', 'config.jsonc'),
      JSON.stringify({
        auto_start_overlay: true,
        mpv: {
          socketPath,
          autoStartSubMiner: true,
          pauseUntilOverlayReady: true,
        },
      }),
    );
    fs.writeFileSync(
      appPath,
      '#!/bin/sh\nif [ -n "$SUBMINER_TEST_CAPTURE" ]; then printf "%s\\n" "$@" > "$SUBMINER_TEST_CAPTURE"; fi\nexit 0\n',
    );
    fs.chmodSync(appPath, 0o755);

    fs.writeFileSync(
      path.join(binDir, 'mpv'),
      `#!/bin/sh
set -eu
printf '%s\\n' "$@" > "$SUBMINER_TEST_MPV_ARGS"
socket_path=""
for arg in "$@"; do
  case "$arg" in
    --input-ipc-server=*)
      socket_path="\${arg#--input-ipc-server=}"
      ;;
  esac
done
${bunBinary} -e "const net=require('node:net'); const fs=require('node:fs'); const path=require('node:path'); const socket=process.argv[1]||''; try{ if (socket) fs.mkdirSync(path.dirname(socket),{recursive:true}); }catch{} try{ if (socket) fs.rmSync(socket,{force:true}); }catch{} if(!socket) process.exit(0); const server=net.createServer((c)=>c.end()); server.on('error',()=>process.exit(0)); try{ server.listen(socket,()=>setTimeout(()=>server.close(()=>process.exit(0)),250)); } catch { process.exit(0); }" "$socket_path"
`,
      'utf8',
    );
    fs.chmodSync(path.join(binDir, 'mpv'), 0o755);
    fs.writeFileSync(path.join(binDir, 'yt-dlp'), '#!/bin/sh\nexit 0\n', 'utf8');
    fs.writeFileSync(path.join(binDir, 'ffmpeg'), '#!/bin/sh\nexit 0\n', 'utf8');
    fs.chmodSync(path.join(binDir, 'yt-dlp'), 0o755);
    fs.chmodSync(path.join(binDir, 'ffmpeg'), 0o755);

    const env = {
      ...makeTestEnv(homeDir, xdgConfigHome),
      PATH: `${binDir}${path.delimiter}${process.env.Path || process.env.PATH || ''}`,
      Path: `${binDir}${path.delimiter}${process.env.Path || process.env.PATH || ''}`,
      DISPLAY: ':99',
      XDG_SESSION_TYPE: 'x11',
      SUBMINER_APPIMAGE_PATH: appPath,
      SUBMINER_TEST_MPV_ARGS: mpvArgsPath,
      SUBMINER_TEST_CAPTURE: path.join(root, 'captured-args.txt'),
    };
    const result = runLauncher(['https://www.youtube.com/watch?v=abc123'], env);

    assert.equal(result.status, 0, `stdout:\n${result.stdout}\nstderr:\n${result.stderr}`);
    const forwardedArgs = fs
      .readFileSync(mpvArgsPath, 'utf8')
      .trim()
      .split('\n')
      .map((item) => item.trim())
      .filter(Boolean);
    assert.equal(forwardedArgs.includes('https://www.youtube.com/watch?v=abc123'), true);
  });
});

test('dictionary command forwards --dictionary and --dictionary-target to app command path', () => {
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
    const targetPath = path.join(root, 'anime-folder');
    fs.mkdirSync(targetPath, { recursive: true });
    const result = runLauncher(['dictionary', targetPath], env);

    assert.equal(result.status, 0);
    assert.equal(
      fs.readFileSync(capturePath, 'utf8'),
      `--start\n--dictionary\n--dictionary-target\n${targetPath}\n`,
    );
  });
});

test('dictionary command forwards manual AniList selection modes to app command path', () => {
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
    const targetPath = path.join(root, 'anime.mkv');
    fs.writeFileSync(targetPath, '');

    assert.equal(runLauncher(['dictionary', '--candidates', targetPath], env).status, 0);
    assert.equal(
      fs.readFileSync(capturePath, 'utf8'),
      `--start\n--dictionary-candidates\n--dictionary-target\n${targetPath}\n`,
    );
    assert.equal(runLauncher(['dictionary', '--select', '21355', targetPath], env).status, 0);
    assert.equal(
      fs.readFileSync(capturePath, 'utf8'),
      `--start\n--dictionary-select\n--dictionary-anilist-id\n21355\n--dictionary-target\n${targetPath}\n`,
    );
  });
});

test(
  'stats command launches attached app flow and waits for response file',
  { timeout: 15000 },
  () => {
    withTempDir((root) => {
      const homeDir = path.join(root, 'home');
      const xdgConfigHome = path.join(root, 'xdg');
      const appPath = path.join(root, 'fake-subminer.sh');
      const capturePath = path.join(root, 'captured-args.txt');
      fs.writeFileSync(
        appPath,
        `#!/bin/sh
set -eu
response_path=""
prev=""
for arg in "$@"; do
  if [ "$prev" = "--stats-response-path" ]; then
    response_path="$arg"
    prev=""
    continue
  fi
  case "$arg" in
    --stats-response-path=*)
      response_path="\${arg#--stats-response-path=}"
      ;;
    --stats-response-path)
      prev="--stats-response-path"
      ;;
  esac
done
if [ -n "$SUBMINER_TEST_STATS_CAPTURE" ]; then
  printf '%s\\n' "$@" > "$SUBMINER_TEST_STATS_CAPTURE"
fi
mkdir -p "$(dirname "$response_path")"
printf '%s' '{"ok":true,"url":"http://127.0.0.1:5175"}' > "$response_path"
exit 0
`,
      );
      fs.chmodSync(appPath, 0o755);

      const env = {
        ...makeTestEnv(homeDir, xdgConfigHome),
        SUBMINER_APPIMAGE_PATH: appPath,
        SUBMINER_TEST_STATS_CAPTURE: capturePath,
      };
      const result = runLauncher(['stats', '--log-level', 'debug'], env);

      assert.equal(result.status, 0, `stdout:\n${result.stdout}\nstderr:\n${result.stderr}`);
      assert.match(
        fs.readFileSync(capturePath, 'utf8'),
        /^--stats\n--stats-response-path\n.+\n--log-level\ndebug\n$/,
      );
    });
  },
);

test(
  'stats command tolerates slower dashboard startup before timing out',
  { timeout: 20000 },
  () => {
    withTempDir((root) => {
      const homeDir = path.join(root, 'home');
      const xdgConfigHome = path.join(root, 'xdg');
      const appPath = path.join(root, 'fake-subminer-slow.sh');
      fs.writeFileSync(
        appPath,
        `#!/bin/sh
set -eu
response_path=""
prev=""
for arg in "$@"; do
  if [ "$prev" = "--stats-response-path" ]; then
    response_path="$arg"
    prev=""
    continue
  fi
  case "$arg" in
    --stats-response-path=*)
      response_path="\${arg#--stats-response-path=}"
      ;;
    --stats-response-path)
      prev="--stats-response-path"
      ;;
  esac
done
sleep 9
mkdir -p "$(dirname "$response_path")"
printf '%s' '{"ok":true,"url":"http://127.0.0.1:5175"}' > "$response_path"
exit 0
`,
      );
      fs.chmodSync(appPath, 0o755);

      const env = {
        ...makeTestEnv(homeDir, xdgConfigHome),
        SUBMINER_APPIMAGE_PATH: appPath,
      };
      const result = runLauncher(['stats'], env);

      assert.equal(result.status, 0, `stdout:\n${result.stdout}\nstderr:\n${result.stderr}`);
    });
  },
);

test('jellyfin discovery routes to app --background and remote announce with log-level forwarding', () => {
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
    assert.equal(
      fs.readFileSync(capturePath, 'utf8'),
      '--background\n--jellyfin-remote-announce\n--log-level\ndebug\n',
    );
  });
});

test('jellyfin discovery via jf alias forwards remote announce for cast visibility', () => {
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
    const result = runLauncher(['-R', 'jf', '--discovery', '--log-level', 'debug'], env);

    assert.equal(result.status, 0);
    assert.equal(
      fs.readFileSync(capturePath, 'utf8'),
      '--background\n--jellyfin-remote-announce\n--log-level\ndebug\n',
    );
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

test('jellyfin setup forwards password-store to app command', () => {
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
    const result = runLauncher(['jf', 'setup', '--password-store', 'gnome-libsecret'], env);

    assert.equal(result.status, 0);
    assert.equal(
      fs.readFileSync(capturePath, 'utf8'),
      '--jellyfin\n--password-store\ngnome-libsecret\n',
    );
  });
});

test('parseJellyfinLibrariesFromAppOutput parses prefixed library lines', () => {
  const parsed = parseJellyfinLibrariesFromAppOutput(`
[subminer] - 2026-03-01 13:10:34 - INFO - [main] Jellyfin library: Anime [lib1] (tvshows)
[subminer] - 2026-03-01 13:10:35 - INFO - [main] Jellyfin library: Movies [lib2] (movies)
`);

  assert.deepEqual(parsed, [
    { id: 'lib1', name: 'Anime', kind: 'tvshows' },
    { id: 'lib2', name: 'Movies', kind: 'movies' },
  ]);
});

test('parseJellyfinItemsFromAppOutput parses item title/id/type tuples', () => {
  const parsed = parseJellyfinItemsFromAppOutput(`
[subminer] - 2026-03-01 13:10:34 - INFO - [main] Jellyfin item: Solo Leveling S01E10 [item-10] (Episode)
[subminer] - 2026-03-01 13:10:35 - INFO - [main] Jellyfin item: Movie [Alt] [movie-1] (Movie)
`);

  assert.deepEqual(parsed, [
    {
      id: 'item-10',
      name: 'Solo Leveling S01E10',
      type: 'Episode',
      display: 'Solo Leveling S01E10',
    },
    {
      id: 'movie-1',
      name: 'Movie [Alt]',
      type: 'Movie',
      display: 'Movie [Alt]',
    },
  ]);
});

test('buildForwardedJellyfinAppArgs forces app log level for parseable list output', () => {
  const forwarded = buildForwardedJellyfinAppArgs(
    {
      jellyfinServer: 'https://jf.example.test/',
      passwordStore: 'gnome-libsecret',
      logLevel: 'info',
    } as never,
    ['--jellyfin-libraries'],
  );

  assert.deepEqual(forwarded, [
    '--jellyfin-libraries',
    '--jellyfin-server',
    'https://jf.example.test',
    '--password-store',
    'gnome-libsecret',
    '--log-level',
    'info',
  ]);
});

test('parseJellyfinErrorFromAppOutput extracts bracketed error lines', () => {
  const parsed = parseJellyfinErrorFromAppOutput(`
[subminer] - 2026-03-01 13:10:34 - WARN - [main] test warning
[2026-03-01T21:11:28.821Z] [ERROR] Missing Jellyfin session. Set SUBMINER_JELLYFIN_ACCESS_TOKEN and SUBMINER_JELLYFIN_USER_ID, then retry.
`);

  assert.equal(
    parsed,
    'Missing Jellyfin session. Set SUBMINER_JELLYFIN_ACCESS_TOKEN and SUBMINER_JELLYFIN_USER_ID, then retry.',
  );
});

test('parseJellyfinErrorFromAppOutput extracts main runtime error lines', () => {
  const parsed = parseJellyfinErrorFromAppOutput(`
[subminer] - 2026-03-01 13:10:34 - ERROR - [main] runJellyfinCommand failed: {"message":"Missing Jellyfin password."}
`);

  assert.equal(
    parsed,
    '[main] runJellyfinCommand failed: {"message":"Missing Jellyfin password."}',
  );
});

test('parseJellyfinPreviewAuthResponse parses valid structured response payload', () => {
  const parsed = parseJellyfinPreviewAuthResponse(
    JSON.stringify({
      serverUrl: 'http://pve-main:8096/',
      accessToken: 'token-123',
      userId: 'user-1',
    }),
  );

  assert.deepEqual(parsed, {
    serverUrl: 'http://pve-main:8096',
    accessToken: 'token-123',
    userId: 'user-1',
  });
});

test('parseJellyfinPreviewAuthResponse returns null for invalid payloads', () => {
  assert.equal(parseJellyfinPreviewAuthResponse(''), null);
  assert.equal(parseJellyfinPreviewAuthResponse('{not json}'), null);
  assert.equal(
    parseJellyfinPreviewAuthResponse(
      JSON.stringify({
        serverUrl: 'http://pve-main:8096',
        accessToken: '',
        userId: 'user-1',
      }),
    ),
    null,
  );
});

test('deriveJellyfinTokenStorePath resolves alongside config path', () => {
  const configPath = path.join('/home/test', '.config', 'SubMiner', 'config.jsonc');
  const tokenPath = deriveJellyfinTokenStorePath(configPath);
  assert.equal(tokenPath, path.join(path.dirname(configPath), 'jellyfin-token-store.json'));
});

test('hasStoredJellyfinSession checks token-store existence', () => {
  const configPath = path.join('/home/test', '.config', 'SubMiner', 'config.jsonc');
  const tokenPath = deriveJellyfinTokenStorePath(configPath);
  const exists = (candidate: string): boolean => candidate === tokenPath;
  assert.equal(hasStoredJellyfinSession(configPath, exists), true);
  assert.equal(
    hasStoredJellyfinSession(path.join('/home/test', '.config', 'Other', 'alt.jsonc'), exists),
    false,
  );
});

test('shouldRetryWithStartForNoRunningInstance matches expected app lifecycle error', () => {
  assert.equal(
    shouldRetryWithStartForNoRunningInstance('No running instance. Use --start to launch the app.'),
    true,
  );
  assert.equal(
    shouldRetryWithStartForNoRunningInstance(
      'Missing Jellyfin session. Run --jellyfin-login first.',
    ),
    false,
  );
});

test('readUtf8FileAppendedSince treats offset as bytes and survives multibyte logs', () => {
  withTempDir((root) => {
    const logPath = path.join(root, 'SubMiner.log');
    const prefix = '[subminer] こんにちは\n';
    const suffix = '[subminer] Jellyfin library: Movies [lib2] (movies)\n';
    fs.writeFileSync(logPath, `${prefix}${suffix}`, 'utf8');

    const byteOffset = Buffer.byteLength(prefix, 'utf8');
    const fromByteOffset = readUtf8FileAppendedSince(logPath, byteOffset);
    assert.match(fromByteOffset, /Jellyfin library: Movies \[lib2\] \(movies\)/);

    const fromBeyondEnd = readUtf8FileAppendedSince(logPath, byteOffset + 9999);
    assert.match(fromBeyondEnd, /Jellyfin library: Movies \[lib2\] \(movies\)/);
  });
});

test('parseEpisodePathFromDisplay extracts series and season from episode display titles', () => {
  assert.deepEqual(
    parseEpisodePathFromDisplay('KONOSUBA S01E03 A Panty Treasure in This Right Hand!'),
    {
      seriesName: 'KONOSUBA',
      seasonNumber: 1,
    },
  );
  assert.deepEqual(parseEpisodePathFromDisplay('Frieren S2E10 Something'), {
    seriesName: 'Frieren',
    seasonNumber: 2,
  });
});

test('parseEpisodePathFromDisplay returns null for non-episode displays', () => {
  assert.equal(parseEpisodePathFromDisplay('Movie Title (Movie)'), null);
  assert.equal(parseEpisodePathFromDisplay('Just A Name'), null);
});

test('buildRootSearchGroups excludes episodes and keeps containers/movies', () => {
  const groups = buildRootSearchGroups([
    { id: 'series-1', name: 'The Eminence in Shadow', type: 'Series', display: 'x' },
    { id: 'movie-1', name: 'Spirited Away', type: 'Movie', display: 'x' },
    { id: 'episode-1', name: 'The Eminence in Shadow S01E01', type: 'Episode', display: 'x' },
  ]);

  assert.deepEqual(groups, [
    {
      id: 'series-1',
      name: 'The Eminence in Shadow',
      type: 'Series',
      display: 'The Eminence in Shadow (Series)',
    },
    {
      id: 'movie-1',
      name: 'Spirited Away',
      type: 'Movie',
      display: 'Spirited Away (Movie)',
    },
  ]);
});

test('classifyJellyfinChildSelection keeps container drilldown state instead of flattening', () => {
  const next = classifyJellyfinChildSelection({ id: 'season-2', type: 'Season' });
  assert.deepEqual(next, {
    kind: 'container',
    id: 'season-2',
  });
});
