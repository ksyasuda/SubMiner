import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';

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
    // Pin the data dir under the temp home so the Linux runtime-plugin preflight
    // resolves managed asset paths deterministically (not the CI runner's).
    XDG_DATA_HOME: path.join(homeDir, '.local', 'share'),
    PATH: pathValue,
    Path: pathValue,
  };
}

// On Linux the playback path runs `ensureLinuxRuntimePluginAvailable`, which
// spawns the app with `--ensure-linux-runtime-plugin-assets` when managed
// support assets are missing and polls up to 30s
// (RESPONSE_TIMEOUT_MS) for an install response. A fake app that just exits
// never writes that response, so the launcher hangs and the test times out on
// Linux CI (the preflight is a no-op on macOS/Windows). This shell prelude makes
// the fake app install the managed support assets and write the response, matching
// launcher/smoke.e2e.test.ts. Prepend it to each fake app that reaches playback.
const RUNTIME_PLUGIN_PREFLIGHT_SH = `if [ "$1" = "--ensure-linux-runtime-plugin-assets" ]; then
  data="\${XDG_DATA_HOME:-$HOME/.local/share}/SubMiner"
  mkdir -p "$data/plugin/subminer" "$data/themes" "$data/thumbnailers"
  printf -- '-- test plugin\\n' > "$data/plugin/subminer/main.lua"
  printf 'test=true\\n' > "$data/plugin/subminer.conf"
  printf '/* test theme */\\n' > "$data/themes/subminer.rasi"
  printf '[Thumbnailer Entry]\\n' > "$data/thumbnailers/subminer-ffmpegthumbnailer.thumbnailer"
  if [ "$2" = "--ensure-linux-runtime-plugin-assets-response-path" ] && [ -n "$3" ]; then
    mkdir -p "$(dirname "$3")"
    printf '{"ok":true,"status":"installed","path":"%s"}' "$data/plugin/subminer/main.lua" > "$3"
  fi
  exit 0
fi
`;

interface Sandbox {
  root: string;
  homeDir: string;
  xdgConfigHome: string;
  /** `<xdgConfigHome>/SubMiner`, created on demand by `writeConfigFile`. */
  configDir: string;
  /** Base launcher env (isolated HOME/XDG); add `SUBMINER_APPIMAGE_PATH` etc. per test. */
  env: NodeJS.ProcessEnv;
  writeConfigFile(name: string, content: string | object): void;
}

/** Runs `fn` inside a temp root with an isolated home and XDG config dir. */
function withSandbox<T>(fn: (sandbox: Sandbox) => T): T {
  return withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const configDir = path.join(xdgConfigHome, 'SubMiner');
    return fn({
      root,
      homeDir,
      xdgConfigHome,
      configDir,
      env: makeTestEnv(homeDir, xdgConfigHome),
      writeConfigFile: (name, content) => {
        fs.mkdirSync(configDir, { recursive: true });
        fs.writeFileSync(
          path.join(configDir, name),
          typeof content === 'string' ? content : JSON.stringify(content),
        );
      },
    });
  });
}

function writeExecutable(filePath: string, content: string): void {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, content, 'utf8');
  fs.chmodSync(filePath, 0o755);
}

/**
 * Writes a fake SubMiner app script. `body` runs after the runtime-plugin preflight prelude
 * (omit the prelude with `preflight: false` for apps that never reach playback). The default
 * body is a no-op; the launcher itself captures forwarded app args when
 * `SUBMINER_TEST_CAPTURE` is set.
 */
function makeFakeApp(
  root: string,
  options: { name?: string; body?: string; preflight?: boolean } = {},
): string {
  const appPath = path.join(root, options.name ?? 'fake-subminer.sh');
  const prelude = options.preflight === false ? '' : RUNTIME_PLUGIN_PREFLIGHT_SH;
  writeExecutable(appPath, `#!/bin/sh\n${prelude}${options.body ?? 'exit 0\n'}`);
  return appPath;
}

/**
 * Fake mpv/yt-dlp/ffmpeg in `<root>/bin`. The mpv stub records its argv into
 * `SUBMINER_TEST_MPV_ARGS` and briefly serves the requested IPC socket so startup succeeds.
 */
function makeFakeMpv(root: string): {
  binDir: string;
  mpvArgsPath: string;
  env: NodeJS.ProcessEnv;
} {
  const binDir = path.join(root, 'bin');
  const mpvArgsPath = path.join(root, 'mpv-args.txt');
  const bunBinary = JSON.stringify(process.execPath.replace(/\\/g, '/'));
  writeExecutable(
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
  );
  writeExecutable(path.join(binDir, 'yt-dlp'), '#!/bin/sh\nexit 0\n');
  writeExecutable(path.join(binDir, 'ffmpeg'), '#!/bin/sh\nexit 0\n');
  const pathValue = `${binDir}${path.delimiter}${process.env.Path || process.env.PATH || ''}`;
  return {
    binDir,
    mpvArgsPath,
    env: { PATH: pathValue, Path: pathValue, SUBMINER_TEST_MPV_ARGS: mpvArgsPath },
  };
}

/**
 * Sets up a completed-setup sandbox that reaches real playback: fake app + fake mpv, a config
 * pointing mpv at a temp socket. Note: no `SUBMINER_TEST_CAPTURE` here. When set, the launcher
 * intercepts every app command, including the Linux runtime-plugin preflight install, and
 * returns without running the fake app, so the preflight would poll 30s for a response.
 */
function setupPlayback(sandbox: Sandbox, options: { autoStart: boolean; extraConfig?: object }) {
  const socketPath = path.join(sandbox.root, 'mpv.sock');
  sandbox.writeConfigFile('setup-state.json', {
    version: 1,
    status: 'completed',
    completedAt: '2026-03-08T00:00:00.000Z',
    completionSource: 'user',
    lastSeenYomitanDictionaryCount: 0,
    pluginInstallStatus: 'installed',
    pluginInstallPathSummary: null,
  });
  sandbox.writeConfigFile('config.jsonc', {
    auto_start_overlay: options.autoStart,
    mpv: {
      socketPath,
      autoStartSubMiner: options.autoStart,
      pauseUntilOverlayReady: options.autoStart,
    },
    ...options.extraConfig,
  });
  const appPath = makeFakeApp(sandbox.root);
  const mpv = makeFakeMpv(sandbox.root);
  const videoPath = path.join(sandbox.root, 'movie.mkv');
  fs.writeFileSync(videoPath, 'fake video content');

  return {
    videoPath,
    env: { ...sandbox.env, ...mpv.env, SUBMINER_APPIMAGE_PATH: appPath },
    readMpvArgs: () => fs.readFileSync(mpv.mpvArgsPath, 'utf8'),
  };
}

function assertSucceeded(result: RunResult): void {
  assert.equal(result.status, 0, `stdout:\n${result.stdout}\nstderr:\n${result.stderr}`);
}

/** Script for a fake app that answers `--stats-response-path` after `sleepSeconds`. */
function statsResponderScript(
  options: { sleepSeconds?: number; captureEnv?: string } = {},
): string {
  return `set -eu
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
${options.captureEnv ? `if [ -n "$${options.captureEnv}" ]; then\n  printf '%s\\n' "$@" > "$${options.captureEnv}"\nfi\n` : ''}${options.sleepSeconds ? `sleep ${options.sleepSeconds}\n` : ''}mkdir -p "$(dirname "$response_path")"
printf '%s' '{"ok":true,"url":"http://127.0.0.1:5175"}' > "$response_path"
exit 0
`;
}

test('config path uses XDG_CONFIG_HOME override', () => {
  withSandbox((sandbox) => {
    sandbox.writeConfigFile('config.json', '{"source":"xdg"}');

    const result = runLauncher(['config', 'path'], sandbox.env);

    assert.equal(result.status, 0);
    assert.equal(result.stdout.trim(), path.join(sandbox.configDir, 'config.json'));
  });
});

for (const flag of ['--version', '-v']) {
  test(`${flag} prints installed app version without requiring app binary`, () => {
    withSandbox((sandbox) => {
      const result = runLauncher([flag], sandbox.env);

      assert.equal(result.status, 0);
      assert.match(result.stdout.trim(), /^SubMiner \d+\.\d+\.\d+/);
      assert.equal(result.stderr, '');
    });
  });
}

test('logs export writes sanitized archive without requiring app binary', () => {
  withSandbox((sandbox) => {
    const logsDir =
      process.platform === 'win32'
        ? path.join(sandbox.xdgConfigHome, 'SubMiner', 'logs')
        : path.join(sandbox.homeDir, '.config', 'SubMiner', 'logs');
    fs.mkdirSync(logsDir, { recursive: true });
    fs.writeFileSync(path.join(logsDir, 'app-2026-W21.log'), `/home/kyle/video.mkv\n`, 'utf8');

    const result = runLauncher(['logs', '-e'], sandbox.env);

    assertSucceeded(result);
    const zipPath = result.stdout.trim();
    assert.match(zipPath, /subminer-logs-.+\.zip$/);
    assert.equal(fs.existsSync(zipPath), true);
    const archive = fs.readFileSync(zipPath);
    assert.equal(archive.includes(Buffer.from('/home/kyle')), false);
    assert.equal(archive.includes(Buffer.from('/home/<user>')), true);
  });
});

test('config path prefers jsonc over json for same directory', () => {
  withSandbox((sandbox) => {
    sandbox.writeConfigFile('config.json', '{"format":"json"}');
    sandbox.writeConfigFile('config.jsonc', '{"format":"jsonc"}');

    const result = runLauncher(['config', 'path'], sandbox.env);

    assert.equal(result.status, 0);
    assert.equal(result.stdout.trim(), path.join(sandbox.configDir, 'config.jsonc'));
  });
});

test('config show prints config body and appends trailing newline', () => {
  withSandbox((sandbox) => {
    sandbox.writeConfigFile('config.jsonc', '{"logLevel":"debug"}');

    const result = runLauncher(['config', 'show'], sandbox.env);

    assert.equal(result.status, 0);
    assert.equal(result.stdout, '{"logLevel":"debug"}\n');
  });
});

test('mpv socket command returns socket path from plugin runtime config', () => {
  withSandbox((sandbox) => {
    const expectedSocket = path.join(sandbox.root, 'custom', 'subminer.sock');
    sandbox.writeConfigFile('config.jsonc', { mpv: { socketPath: expectedSocket } });

    const result = runLauncher(['mpv', 'socket'], sandbox.env);

    assert.equal(result.status, 0);
    assert.equal(result.stdout.trim(), expectedSocket);
  });
});

test('mpv status exits non-zero when socket is not ready', () => {
  withSandbox((sandbox) => {
    sandbox.writeConfigFile('config.jsonc', {
      mpv: { socketPath: path.join(sandbox.root, 'missing.sock') },
    });

    const result = runLauncher(['mpv', 'status'], sandbox.env);

    assert.equal(result.status, 1);
    assert.match(result.stdout, /socket not ready/i);
  });
});

test('doctor reports checks and exits non-zero without hard dependencies', () => {
  withSandbox((sandbox) => {
    const result = runLauncher(['doctor'], { ...sandbox.env, PATH: '', Path: '' });

    assert.equal(result.status, 1);
    assert.match(result.stdout, /\[doctor\] app binary:/);
    assert.match(result.stdout, /\[doctor\] mpv:/);
    assert.match(result.stdout, /\[doctor\] config:/);
  });
});

test('doctor refresh-known-words forwards app refresh command without requiring mpv', () => {
  withSandbox((sandbox) => {
    const capturePath = path.join(sandbox.root, 'captured-args.txt');
    const env = {
      ...sandbox.env,
      PATH: '',
      Path: '',
      SUBMINER_APPIMAGE_PATH: makeFakeApp(sandbox.root),
      SUBMINER_TEST_CAPTURE: capturePath,
    };

    const result = runLauncher(['doctor', '--refresh-known-words'], env);

    assert.equal(result.status, 0);
    assert.equal(fs.readFileSync(capturePath, 'utf8'), '--refresh-known-words\n');
    assert.match(result.stdout, /\[doctor\] mpv: missing/);
  });
});

const forwardedAppArgsCases: Array<{
  name: string;
  /** Receives the sandbox root so cases can create targets (dictionary folders). */
  argv: (root: string) => string[];
  expected: (root: string) => string;
}> = [
  {
    name: 'launcher settings command forwards app settings window command',
    argv: () => ['settings'],
    expected: () => '--settings\n',
  },
  {
    name: 'launcher youtube command forwards the YouTube browser flag and log level',
    argv: () => ['yt', '--log-level', 'debug'],
    expected: () => '--youtube-browser\n--log-level\ndebug\n',
  },
  {
    name: 'dictionary command forwards --dictionary and --dictionary-target to app command path',
    argv: (root) => ['dictionary', path.join(root, 'anime-folder')],
    expected: (root) =>
      `--start\n--dictionary\n--dictionary-target\n${path.join(root, 'anime-folder')}\n`,
  },
  {
    name: 'jellyfin discovery routes to app --background and remote announce with log-level forwarding',
    argv: () => ['jellyfin', 'discovery', '--log-level', 'debug'],
    expected: () => '--background\n--jellyfin-remote-announce\n--log-level\ndebug\n',
  },
  {
    name: 'jellyfin discovery via jf alias forwards remote announce for cast visibility',
    argv: () => ['-R', 'jf', '--discovery', '--log-level', 'debug'],
    expected: () => '--background\n--jellyfin-remote-announce\n--log-level\ndebug\n',
  },
  {
    name: 'jellyfin login routes credentials to app command',
    argv: () => [
      'jellyfin',
      'login',
      '--server',
      'https://jf.example.test',
      '--username',
      'alice',
      '--password',
      'secret',
    ],
    expected: () =>
      '--jellyfin-login\n--jellyfin-server\nhttps://jf.example.test\n--jellyfin-username\nalice\n--jellyfin-password\nsecret\n',
  },
  {
    name: 'jellyfin setup forwards password-store to app command',
    argv: () => ['jf', 'setup', '--password-store', 'gnome-libsecret'],
    expected: () => '--jellyfin\n--password-store\ngnome-libsecret\n',
  },
];

for (const c of forwardedAppArgsCases) {
  test(c.name, () => {
    withSandbox((sandbox) => {
      fs.mkdirSync(path.join(sandbox.root, 'anime-folder'), { recursive: true });
      const capturePath = path.join(sandbox.root, 'captured-args.txt');
      const env = {
        ...sandbox.env,
        SUBMINER_APPIMAGE_PATH: makeFakeApp(sandbox.root),
        SUBMINER_TEST_CAPTURE: capturePath,
      };

      const result = runLauncher(c.argv(sandbox.root), env);

      assert.equal(result.status, 0);
      assert.equal(fs.readFileSync(capturePath, 'utf8'), c.expected(sandbox.root));
    });
  });
}

test('launcher settings command suppresses known Electron macOS menu diagnostics', () => {
  withSandbox((sandbox) => {
    const appPath = makeFakeApp(sandbox.root, {
      preflight: false,
      body: [
        'printf "%s\\n" "2026-05-17 02:59:52.141 SubMiner[29060:305323] representedObject is not a WeakPtrToElectronMenuModelAsNSObject" >&2',
        'printf "%s\\n" "real stderr line" >&2',
        'exit 0',
        '',
      ].join('\n'),
    });

    const result = runLauncher(['settings'], { ...sandbox.env, SUBMINER_APPIMAGE_PATH: appPath });

    assert.equal(result.status, 0);
    assert.equal(result.stderr, 'real stderr line\n');
  });
});

test('launcher forwards --args to mpv as parsed tokens', { timeout: 15000 }, () => {
  withSandbox((sandbox) => {
    const playback = setupPlayback(sandbox, { autoStart: false });

    const result = runLauncher(
      ['--args', '--pause=yes --title="movie night"', playback.videoPath],
      playback.env,
    );

    assertSucceeded(result);
    const forwardedArgs = playback
      .readMpvArgs()
      .trim()
      .split('\n')
      .map((item) => item.trim())
      .filter(Boolean);
    assert.equal(forwardedArgs.includes('--pause=yes'), true);
    assert.equal(forwardedArgs.includes('--title=movie night'), true);
    assert.equal(forwardedArgs.includes(playback.videoPath), true);
  });
});

test('launcher forwards non-info log level into mpv logging args', { timeout: 15000 }, () => {
  withSandbox((sandbox) => {
    const playback = setupPlayback(sandbox, {
      autoStart: true,
      extraConfig: { logging: { files: { mpv: true } } },
    });

    const result = runLauncher(['--log-level', 'debug', playback.videoPath], playback.env);

    assertSucceeded(result);
    const mpvArgs = playback.readMpvArgs();
    assert.match(mpvArgs, /--msg-level=all=warn,subminer=debug/);
    assert.doesNotMatch(mpvArgs, /--script-opts=.*subminer-log_level=debug/);
  });
});

test('launcher routes youtube urls through regular playback startup', { timeout: 15000 }, () => {
  withSandbox((sandbox) => {
    const playback = setupPlayback(sandbox, { autoStart: true });
    const url = 'https://www.youtube.com/watch?v=abc123';

    // Pass an explicit backend so overlay startup doesn't probe for a display
    // (headless CI has none), matching launcher/smoke.e2e.test.ts.
    const result = runLauncher(['--backend', 'x11', url], playback.env);

    assertSucceeded(result);
    const forwardedArgs = playback
      .readMpvArgs()
      .trim()
      .split('\n')
      .map((item) => item.trim())
      .filter(Boolean);
    assert.equal(forwardedArgs.includes(url), true);
  });
});

test(
  'stats command launches attached app flow and waits for response file',
  { timeout: 15000 },
  () => {
    withSandbox((sandbox) => {
      const capturePath = path.join(sandbox.root, 'captured-args.txt');
      const appPath = makeFakeApp(sandbox.root, {
        preflight: false,
        body: statsResponderScript({ captureEnv: 'SUBMINER_TEST_STATS_CAPTURE' }),
      });
      const env = {
        ...sandbox.env,
        SUBMINER_APPIMAGE_PATH: appPath,
        SUBMINER_TEST_STATS_CAPTURE: capturePath,
      };

      const result = runLauncher(['stats', '--log-level', 'debug'], env);

      assertSucceeded(result);
      assert.match(
        fs.readFileSync(capturePath, 'utf8'),
        /^--stats\n--stats-response-path\n.+\n--log-level\ndebug\n$/,
      );
    });
  },
);

// The startup response timeout (12s) is not overridable, so this really waits 9s.
test(
  'stats command tolerates slower dashboard startup before timing out',
  { timeout: 20000 },
  () => {
    withSandbox((sandbox) => {
      const appPath = makeFakeApp(sandbox.root, {
        name: 'fake-subminer-slow.sh',
        preflight: false,
        body: statsResponderScript({ sleepSeconds: 9 }),
      });

      const result = runLauncher(['stats'], { ...sandbox.env, SUBMINER_APPIMAGE_PATH: appPath });

      assertSucceeded(result);
    });
  },
);
