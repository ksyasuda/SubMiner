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
} from './jellyfin.js';

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
  const result = spawnSync(
    process.execPath,
    ['run', path.join(process.cwd(), 'launcher/main.ts'), ...argv],
    {
      env,
      encoding: 'utf8',
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
    const socketPath = path.join(root, 'missing.sock');
    fs.mkdirSync(path.join(xdgConfigHome, 'mpv', 'script-opts'), { recursive: true });
    fs.writeFileSync(
      path.join(xdgConfigHome, 'mpv', 'script-opts', 'subminer.conf'),
      `socket_path=${socketPath}\n`,
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

test('youtube command rejects removed --mode option', () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const appPath = path.join(root, 'fake-subminer.sh');
    fs.writeFileSync(appPath, '#!/bin/sh\nexit 0\n');
    fs.chmodSync(appPath, 0o755);

    const env = {
      ...makeTestEnv(homeDir, xdgConfigHome),
      SUBMINER_APPIMAGE_PATH: appPath,
    };
    const result = runLauncher(
      ['youtube', 'https://www.youtube.com/watch?v=test123', '--mode', 'automatic'],
      env,
    );

    assert.equal(result.status, 1);
    assert.match(result.stderr, /unknown option '--mode'/i);
  });
});

test('youtube playback generates subtitles before mpv launch', { timeout: 15000 }, () => {
  withTempDir((root) => {
    const homeDir = path.join(root, 'home');
    const xdgConfigHome = path.join(root, 'xdg');
    const binDir = path.join(root, 'bin');
    const appPath = path.join(root, 'fake-subminer.sh');
    const ytdlpLogPath = path.join(root, 'yt-dlp.log');
    const mpvCapturePath = path.join(root, 'mpv-order.txt');
    const mpvArgsPath = path.join(root, 'mpv-args.txt');
    const socketPath = path.join(root, 'mpv.sock');
    const bunBinary = JSON.stringify(process.execPath.replace(/\\/g, '/'));

    fs.mkdirSync(binDir, { recursive: true });
    fs.mkdirSync(path.join(xdgConfigHome, 'SubMiner'), { recursive: true });
    fs.mkdirSync(path.join(xdgConfigHome, 'mpv', 'script-opts'), { recursive: true });
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
      path.join(xdgConfigHome, 'mpv', 'script-opts', 'subminer.conf'),
      `socket_path=${socketPath}\nauto_start=no\nauto_start_visible_overlay=no\nauto_start_pause_until_ready=no\n`,
    );
    fs.writeFileSync(appPath, '#!/bin/sh\nexit 0\n');
    fs.chmodSync(appPath, 0o755);

    fs.writeFileSync(
      path.join(binDir, 'yt-dlp'),
      `#!/bin/sh
set -eu
printf '%s\\n' "$*" >> "$SUBMINER_TEST_YTDLP_LOG"
if printf '%s\\n' "$*" | grep -q -- '--dump-single-json'; then
  printf '{"id":"video123"}\\n'
  exit 0
fi
out_dir=""
prev=""
for arg in "$@"; do
  if [ "$prev" = "-o" ]; then
    out_dir=$(dirname "$arg")
    break
  fi
  prev="$arg"
done
mkdir -p "$out_dir"
printf '1\\n00:00:00,000 --> 00:00:01,000\\nこんにちは\\n' > "$out_dir/video123.ja.srt"
printf '1\\n00:00:00,000 --> 00:00:01,000\\nhello\\n' > "$out_dir/video123.en.srt"
`,
      'utf8',
    );
    fs.chmodSync(path.join(binDir, 'yt-dlp'), 0o755);

    fs.writeFileSync(path.join(binDir, 'ffmpeg'), '#!/bin/sh\nexit 0\n', 'utf8');
    fs.chmodSync(path.join(binDir, 'ffmpeg'), 0o755);

    fs.writeFileSync(
      path.join(binDir, 'mpv'),
      `#!/bin/sh
set -eu
if [ -s "$SUBMINER_TEST_YTDLP_LOG" ]; then
  printf 'generated-before-mpv\\n' > "$SUBMINER_TEST_MPV_ORDER"
else
  printf 'mpv-before-generation\\n' > "$SUBMINER_TEST_MPV_ORDER"
fi
printf '%s\\n' "$@" > "$SUBMINER_TEST_MPV_ARGS"
socket_path=""
for arg in "$@"; do
  case "$arg" in
    --input-ipc-server=*)
      socket_path="\${arg#--input-ipc-server=}"
      ;;
  esac
done
${bunBinary} -e "const net=require('node:net'); const fs=require('node:fs'); const socket=process.argv[1]; try { fs.rmSync(socket,{force:true}); } catch {} const server=net.createServer((conn)=>conn.end()); server.listen(socket,()=>setTimeout(()=>server.close(()=>process.exit(0)),250));" "$socket_path"
`,
      'utf8',
    );
    fs.chmodSync(path.join(binDir, 'mpv'), 0o755);

    const env = {
      ...makeTestEnv(homeDir, xdgConfigHome),
      PATH: `${binDir}${path.delimiter}${process.env.Path || process.env.PATH || ''}`,
      Path: `${binDir}${path.delimiter}${process.env.Path || process.env.PATH || ''}`,
      SUBMINER_APPIMAGE_PATH: appPath,
      SUBMINER_TEST_YTDLP_LOG: ytdlpLogPath,
      SUBMINER_TEST_MPV_ORDER: mpvCapturePath,
      SUBMINER_TEST_MPV_ARGS: mpvArgsPath,
    };
    const result = runLauncher(['youtube', 'https://www.youtube.com/watch?v=test123'], env);

    assert.equal(result.status, 0, `stdout:\n${result.stdout}\nstderr:\n${result.stderr}`);
    assert.equal(fs.readFileSync(mpvCapturePath, 'utf8').trim(), 'generated-before-mpv');
    assert.match(
      fs.readFileSync(mpvArgsPath, 'utf8'),
      /https:\/\/www\.youtube\.com\/watch\?v=test123/,
    );
    assert.match(fs.readFileSync(ytdlpLogPath, 'utf8'), /--dump-single-json/);
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
      `--dictionary\n--dictionary-target\n${targetPath}\n`,
    );
  });
});

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
