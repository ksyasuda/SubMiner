import { spawn, type ChildProcess } from 'node:child_process';
import { once } from 'node:events';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { connectCdpPage, listCdpTargets, type CdpPage } from './cdp';
import { startDisplay, type E2eDisplay } from './display';
import { FAKE_ANKI_DECK, FAKE_ANKI_MODEL, startFakeAnki, type FakeAnki } from './fake-anki';
import { buildFixtureDictionary, generateFixtureClip, writeFixtureSubtitles } from './fixtures';
import { connectMpv, startMpv, type MpvClient } from './mpv';
import { applyEnvDelta, stopProcess, type EnvDelta } from './process';
import { waitUntil } from './wait';
import {
  shouldForceX11ElectronBackend,
  X11_ELECTRON_BOOTSTRAP_ENV,
} from '../../src/core/utils/electron-backend';

const repoRoot = path.resolve(__dirname, '..', '..');

/** URL fragments identifying the app's windows among the CDP targets. */
export const OVERLAY_PAGE = 'renderer/index.html';
const YOMITAN_SETTINGS_PAGE = 'settings.html';

type JsonObject = { [key: string]: JsonValue };
type JsonValue = string | number | boolean | null | JsonValue[] | JsonObject;

export type E2eSessionOptions = {
  /** Deep-merged over the harness config, for scenarios that need other settings. */
  config?: JsonObject;
};

export type AppTarget = {
  /** Environment changes that point a process at the isolated instance. */
  env: EnvDelta;
  /** Electron/Chromium switches every launch needs. */
  electronArgs: string[];
  /** Receives stdout and stderr of every launch. */
  logPath: string;
};

export type E2eSession = {
  /** Temp dir holding the isolated config, fixtures, sockets, and logs. */
  root: string;
  display: E2eDisplay['kind'];
  cdpPort: number;
  mpvSocketPath: string;
  /** How to start a process that talks to this isolated instance. */
  target: AppTarget;
  mpv: MpvClient;
  anki: FakeAnki;
  /** Connects to the first window or frame whose URL contains `urlPart`, waiting for it to exist. */
  openPage: (urlPart: string) => Promise<CdpPage>;
  /** Runs a CLI command (e.g. `--mine-sentence`) against the running instance. */
  runAppCommand: (...args: string[]) => Promise<void>;
  dispose: () => Promise<void>;
};

function isJsonObject(value: JsonValue | undefined): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function mergeConfig(base: JsonObject, overrides: JsonObject): JsonObject {
  const merged = { ...base };
  for (const [key, value] of Object.entries(overrides)) {
    const existing = merged[key];
    merged[key] =
      isJsonObject(existing) && isJsonObject(value) ? mergeConfig(existing, value) : value;
  }
  return merged;
}

// Everything that could reach outside the sandbox (network services, the real
// Anki, ports the user's own SubMiner holds) is off or pointed at a fake.
function buildConfig(mpvSocketPath: string, ankiUrl: string): JsonObject {
  return {
    mpv: { socketPath: mpvSocketPath, aniskipEnabled: false },
    ankiConnect: {
      url: ankiUrl,
      deck: FAKE_ANKI_DECK,
      proxy: { enabled: false },
      isLapis: { enabled: true, sentenceCardModel: FAKE_ANKI_MODEL },
    },
    stats: { autoStartServer: false },
    updates: { enabled: false },
    discordPresence: { enabled: false },
    websocket: { enabled: false },
    annotationWebsocket: { enabled: false },
    texthooker: { launchAtStartup: false },
  };
}

// Outside Electron the package's entry point is the path of its binary, and
// resolving it downloads the binary when an install skipped that step (CI).
function resolveElectronBinary(): string {
  const binary: unknown = require('electron');
  if (typeof binary !== 'string') throw new Error('Could not resolve the Electron binary path');
  return binary;
}

function readDevToolsPort(portFile: string): number | null {
  try {
    const port = Number(fs.readFileSync(portFile, 'utf8').split('\n')[0]);
    return port > 0 ? port : null;
  } catch {
    return null;
  }
}

function spawnApp(target: AppTarget, electronArgs: string[], args: string[]): ChildProcess {
  const log = fs.openSync(target.logPath, 'a');
  const child = spawn(
    resolveElectronBinary(),
    [...target.electronArgs, ...electronArgs, repoRoot, ...args],
    { cwd: repoRoot, env: applyEnvDelta(process.env, target.env), stdio: ['ignore', log, log] },
  );
  fs.closeSync(log);
  return child;
}

/** Runs a CLI command against the running instance and waits for it to finish. */
export async function runAppCommand(target: AppTarget, args: string[]): Promise<void> {
  const child = spawnApp(target, [], args);
  const timer = setTimeout(() => child.kill('SIGKILL'), 30_000);
  const [code] = await once(child, 'exit');
  clearTimeout(timer);
  if (code !== 0) {
    throw new Error(
      `SubMiner ${args.join(' ')} exited with ${String(code)}; see ${target.logPath}`,
    );
  }
}

async function openPage(cdpPort: number, urlPart: string): Promise<CdpPage> {
  const target = await waitUntil(
    async () =>
      (await listCdpTargets(cdpPort)).find((candidate) => candidate.url.includes(urlPart)),
    { description: `app window matching "${urlPart}"`, timeoutMs: 30_000 },
  );
  const page = await connectCdpPage(target);
  // A target is listed before its document exists; callers expect a usable DOM.
  await page.waitFor(`document.readyState === "complete"`, {
    description: `"${urlPart}" document to load`,
  });
  return page;
}

export async function startE2eSession(options: E2eSessionOptions = {}): Promise<E2eSession> {
  if (!fs.existsSync(path.join(repoRoot, 'dist', 'main-entry.js'))) {
    throw new Error('dist/ is missing. Run `bun run build` before the e2e harness.');
  }

  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-e2e-'));
  const configHome = path.join(root, 'config');
  const userDataDir = path.join(configHome, 'SubMiner');
  const mediaPath = path.join(root, 'clip.mkv');
  const subtitlePath = path.join(root, 'clip.srt');
  const dictionaryPath = path.join(root, 'fixture-dictionary.zip');
  const mpvSocketPath =
    process.platform === 'win32'
      ? `\\\\.\\pipe\\${path.basename(root)}-mpv`
      : path.join(root, 'mpv.sock');

  // Teardown steps, run in reverse. Also used to unwind a failed startup.
  const teardown: Array<() => Promise<void> | void> = [];
  const dispose = async (): Promise<void> => {
    for (const step of teardown.reverse()) await step();
    teardown.length = 0;
    if (process.env.SUBMINER_E2E_KEEP !== '1') {
      fs.rmSync(root, { recursive: true, force: true });
    }
  };

  try {
    fs.mkdirSync(userDataDir, { recursive: true });
    writeFixtureSubtitles(subtitlePath);
    await Promise.all([generateFixtureClip(mediaPath), buildFixtureDictionary(dictionaryPath)]);

    const display = await startDisplay();
    teardown.push(display.stop);
    const anki = await startFakeAnki();
    teardown.push(anki.stop);

    const config = mergeConfig(buildConfig(mpvSocketPath, anki.url), options.config ?? {});
    fs.writeFileSync(path.join(userDataDir, 'config.jsonc'), JSON.stringify(config, null, 2));

    const env: EnvDelta = {
      set: {
        ...display.env.set,
        // The app derives its config dir, Electron profile, and control socket from these.
        ...(process.platform === 'win32'
          ? { APPDATA: configHome }
          : { XDG_CONFIG_HOME: configHome, XDG_DATA_HOME: path.join(root, 'data') }),
        SUBMINER_APP_LOG: path.join(root, 'app.log'),
        // Stay attached instead of re-spawning detached, so the harness owns the process.
        SUBMINER_BACKGROUND_CHILD: '1',
      },
      unset: [...display.env.unset, 'ELECTRON_RUN_AS_NODE'],
    };
    // --enable-logging routes renderer console output into the app log too.
    const electronArgs = ['--password-store=basic', '--enable-logging=stderr'];
    // Outside Hyprland/Sway the app re-spawns itself detached onto the X11
    // backend. Start it there directly so the process we spawn is the app.
    if (shouldForceX11ElectronBackend(applyEnvDelta(process.env, env))) {
      env.set[X11_ELECTRON_BOOTSTRAP_ENV] = '1';
      electronArgs.push('--ozone-platform=x11');
    }
    // CI containers and runners lack the user namespaces Chromium's sandbox needs.
    if (process.platform === 'linux' && process.env.CI) electronArgs.push('--no-sandbox');
    const target: AppTarget = { env, electronArgs, logPath: path.join(root, 'app-stdio.log') };

    const launchApp = async (args: string[]): Promise<{ child: ChildProcess; cdpPort: number }> => {
      // With port 0 Chromium picks a free port and records it in the profile dir.
      const portFile = path.join(userDataDir, 'DevToolsActivePort');
      fs.rmSync(portFile, { force: true });
      const child = spawnApp(
        target,
        ['--remote-debugging-port=0'],
        [...args, '--log-level', 'debug'],
      );
      const stop = () => stopProcess(child);
      teardown.push(stop);
      const cdpPort = await waitUntil(
        () => {
          if (child.exitCode !== null) {
            throw new Error(`SubMiner exited with code ${child.exitCode} during startup`);
          }
          return readDevToolsPort(portFile);
        },
        { description: 'SubMiner DevTools port', timeoutMs: 30_000 },
      );
      return { child, cdpPort };
    };

    // First boot: import the fixture dictionary through Yomitan's settings page.
    // Without one the tokenizer finds nothing and first-run setup takes over.
    const seeder = await launchApp(['--yomitan']);
    const settings = await openPage(seeder.cdpPort, YOMITAN_SETTINGS_PAGE);
    await settings.waitFor('globalThis.__subminerYomitanSettingsAutomation?.ready === true', {
      description: 'Yomitan settings automation bridge',
    });
    const archive = fs.readFileSync(dictionaryPath).toString('base64');
    await settings.evaluate(
      `globalThis.__subminerYomitanSettingsAutomation
        .importDictionaryArchiveBase64(${JSON.stringify(archive)}, "fixture-dictionary.zip")
        .then(() => true)`,
    );
    settings.close();
    await stopProcess(seeder.child);

    const mpvProcess = await startMpv({
      socketPath: mpvSocketPath,
      mediaPath,
      subtitlePath,
      logPath: path.join(root, 'mpv.log'),
      display,
    });
    teardown.push(() => stopProcess(mpvProcess));
    const mpv = await connectMpv(mpvSocketPath);
    teardown.push(mpv.close);

    const app = await launchApp(['--start']);
    // mpv starts paused and the app resumes it once the overlay and tokenizer
    // are ready, so that resume marks the end of startup.
    await waitUntil(async () => (await mpv.command('get_property', 'pause')) === false, {
      description: 'SubMiner to resume playback after overlay startup',
      timeoutMs: 30_000,
    });
    // Hand scenarios a paused player. The app retries its resume for a moment,
    // so keep pausing until the pause has held for longer than that retry window.
    let pausedSince: number | null = null;
    await waitUntil(
      async () => {
        if ((await mpv.command('get_property', 'pause')) !== true) {
          pausedSince = null;
          await mpv.command('set_property', 'pause', true);
          return false;
        }
        pausedSince ??= Date.now();
        return Date.now() - pausedSince >= 1_500;
      },
      { description: 'playback to stay paused after startup', timeoutMs: 40_000 },
    );
    const pages: CdpPage[] = [];
    teardown.push(() => pages.forEach((page) => page.close()));

    return {
      root,
      display: display.kind,
      cdpPort: app.cdpPort,
      mpvSocketPath,
      target,
      mpv,
      anki,
      openPage: async (urlPart) => {
        const page = await openPage(app.cdpPort, urlPart);
        pages.push(page);
        return page;
      },
      runAppCommand: (...args) => runAppCommand(target, args),
      dispose,
    };
  } catch (error) {
    // Keep the logs: a failed startup is exactly when they are needed.
    process.env.SUBMINER_E2E_KEEP = '1';
    await dispose();
    throw new Error(`E2E session failed to start (logs kept in ${root})`, { cause: error });
  }
}
