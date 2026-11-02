import assert from 'node:assert/strict';
import test from 'node:test';
import { createAppState } from '../state';
import { createPluginsAssembly, type PluginsAssemblyDeps } from './plugins';

type MpvClient = NonNullable<ReturnType<typeof createAppState>['mpvClient']>;

function lazy(name: string) {
  return (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
}

/** Builds deps whose every function throws, so construction must not call any of them. */
function makeDeps(): PluginsAssemblyDeps {
  const thrower = <T extends object>(name: string): T =>
    new Proxy({}, { get: (_target, key: string) => lazy(`${name}.${key}`) }) as T;
  return {
    appState: createAppState({ mpvSocketPath: '/tmp/subminer-plugins.sock', texthookerPort: 0 }),
    configService: thrower('configService'),
    overlayManager: thrower('overlayManager'),
    logger: thrower('logger'),
    createLogger: lazy('createLogger'),
    configDir: '/tmp/subminer-plugins-test',
    userDataPath: '/tmp/subminer-plugins-test',
    overlayModalRuntime: thrower('overlayModalRuntime'),
    overlayModals: thrower('overlayModals'),
    getOverlayNotifications: lazy('getOverlayNotifications'),
    createBrowserWindow: lazy('createBrowserWindow'),
    promoteAboveOverlay: lazy('promoteAboveOverlay'),
    ensureMpvRunning: lazy('ensureMpvRunning'),
    dispatchSessionAction: lazy('dispatchSessionAction'),
  };
}

test('createPluginsAssembly does not call its deps while constructing', () => {
  const plugins = createPluginsAssembly(makeDeps());
  assert.equal(typeof plugins.host.loadAll, 'function');
});

test('plugins see the current line with its parsed cue timing', () => {
  const deps = makeDeps();
  deps.appState.currentSubText = '一行目';
  deps.appState.activeParsedSubtitleCues = [{ startTime: 10, endTime: 12, text: '一行目' }];
  deps.appState.mpvClient = {
    currentSubText: '一行目',
    currentTimePos: 11,
    currentSubDelay: 0,
  } as MpvClient;

  assert.deepEqual(createPluginsAssembly(deps).getCurrentSubtitle(), {
    text: '一行目',
    startTime: 10,
    endTime: 12,
  });
});

test('without parsed cues plugins get the live line, timed only while mpv reports a span', () => {
  const deps = makeDeps();
  const plugins = createPluginsAssembly(deps);
  deps.appState.currentSubText = 'ライブ';
  deps.appState.mpvClient = {
    currentSubText: 'ライブ',
    currentSubStart: 3,
    currentSubEnd: 5,
  } as MpvClient;
  assert.deepEqual(plugins.getCurrentSubtitle(), { text: 'ライブ', startTime: 3, endTime: 5 });

  // A cleared sub-start reads as zero; a span that does not move forward is not a timing.
  deps.appState.mpvClient = {
    currentSubText: 'ライブ',
    currentSubStart: 0,
    currentSubEnd: 0,
  } as MpvClient;
  assert.deepEqual(plugins.getCurrentSubtitle(), {
    text: 'ライブ',
    startTime: null,
    endTime: null,
  });

  deps.appState.currentSubText = '  ';
  assert.equal(plugins.getCurrentSubtitle(), null);
});
