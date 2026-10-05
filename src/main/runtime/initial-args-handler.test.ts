import test from 'node:test';
import assert from 'node:assert/strict';
import { parseArgs, type CliArgs } from '../../cli/args';
import { createHandleInitialArgsHandler } from './initial-args-handler';

type HandlerDeps = Parameters<typeof createHandleInitialArgsHandler>[0];

function cliArgs(overrides: Partial<CliArgs> = {}): CliArgs {
  return { ...parseArgs([]), start: true, ...overrides };
}

// Records every startup side effect so tests can assert what the handler did.
function runInitialArgs(
  overrides: Partial<HandlerDeps> & { mpvClientConnected?: boolean } = {},
): string[] {
  const effects: string[] = [];
  const { mpvClientConnected, ...depOverrides } = overrides;
  const mpvClient =
    mpvClientConnected === undefined
      ? null
      : { connected: mpvClientConnected, connect: () => effects.push('connect') };
  createHandleInitialArgsHandler({
    getInitialArgs: () => cliArgs(),
    isBackgroundMode: () => false,
    shouldEnsureTrayOnStartup: () => false,
    shouldRunHeadlessInitialCommand: () => false,
    ensureTray: () => effects.push('tray'),
    isTexthookerOnlyMode: () => false,
    hasImmersionTracker: () => false,
    getMpvClient: () => mpvClient,
    commandNeedsOverlayStartupPrereqs: () => false,
    commandNeedsOverlayRuntime: () => false,
    ensureOverlayStartupPrereqs: () => effects.push('prereqs'),
    isOverlayRuntimeInitialized: () => false,
    initializeOverlayRuntime: () => effects.push('init-overlay'),
    logInfo: (message) => effects.push(`log:${message}`),
    handleCliCommand: (_args, source) => effects.push(`cli:${source}`),
    ...depOverrides,
  })();
  return effects;
}

const AUTO_CONNECT_LOG = 'log:Auto-connecting MPV client for immersion tracking';

const cases: Array<{
  name: string;
  overrides: Parameters<typeof runInitialArgs>[0];
  expected: string[];
}> = [
  {
    name: 'no-ops without initial args',
    overrides: { getInitialArgs: () => null },
    expected: [],
  },
  {
    name: 'forwards args to cli handler without bootstrapping the overlay',
    overrides: {},
    expected: ['cli:initial'],
  },
  {
    name: 'ensures tray in background mode',
    overrides: { isBackgroundMode: () => true },
    expected: ['tray', 'cli:initial'],
  },
  {
    name: 'can ensure tray outside background mode when requested',
    overrides: { shouldEnsureTrayOnStartup: () => true },
    expected: ['tray', 'cli:initial'],
  },
  {
    name: 'auto-connects mpv for immersion tracking',
    overrides: { hasImmersionTracker: () => true, mpvClientConnected: false },
    expected: [AUTO_CONNECT_LOG, 'connect', 'cli:initial'],
  },
  {
    name: 'skips mpv auto-connect for stats mode',
    overrides: {
      getInitialArgs: () => cliArgs({ start: false, stats: true }),
      hasImmersionTracker: () => true,
      mpvClientConnected: false,
    },
    expected: ['cli:initial'],
  },
  {
    name: 'bootstraps overlay before initial overlay-runtime commands',
    overrides: {
      commandNeedsOverlayStartupPrereqs: () => true,
      commandNeedsOverlayRuntime: () => true,
    },
    expected: ['prereqs', 'init-overlay', 'cli:initial'],
  },
  {
    name: 'prepares prereqs but skips eager overlay bootstrap for youtube playback',
    overrides: {
      getInitialArgs: () => cliArgs({ youtubePlay: 'https://youtube.com/watch?v=abc' }),
      commandNeedsOverlayStartupPrereqs: () => true,
    },
    expected: ['prereqs', 'cli:initial'],
  },
  {
    name: 'skips tray, mpv auto-connect, and overlay bootstrap for headless refresh',
    overrides: {
      getInitialArgs: () => cliArgs({ start: false, refreshKnownWords: true }),
      isBackgroundMode: () => true,
      shouldEnsureTrayOnStartup: () => true,
      shouldRunHeadlessInitialCommand: () => true,
      hasImmersionTracker: () => true,
      mpvClientConnected: false,
      commandNeedsOverlayStartupPrereqs: () => true,
      commandNeedsOverlayRuntime: () => true,
    },
    expected: ['cli:initial'],
  },
];

for (const c of cases) {
  test(`initial args handler ${c.name}`, () => {
    assert.deepEqual(runInitialArgs(c.overrides), c.expected);
  });
}
