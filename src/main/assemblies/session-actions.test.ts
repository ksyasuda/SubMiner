import assert from 'node:assert/strict';
import test from 'node:test';
import { DEFAULT_CONFIG } from '../../config/definitions';
import type { ResolvedConfig } from '../../types';
import { createAppState } from '../state';
import { createSessionActionsAssembly, type SessionActionsAssemblyDeps } from './session-actions';

function lazy(name: string) {
  return (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
}

/** Builds deps whose every function throws unless a test overrides it. */
function makeDeps(overrides: Partial<SessionActionsAssemblyDeps> = {}): SessionActionsAssemblyDeps {
  const thrower = <T extends object>(name: string): T =>
    new Proxy({}, { get: (_target, key: string) => lazy(`${name}.${key}`) }) as T;
  return {
    appState: createAppState({ mpvSocketPath: '/tmp/subminer-session.sock', texthookerPort: 0 }),
    getConfig: lazy('getConfig'),
    logger: thrower('logger'),
    overlayManager: thrower('overlayManager'),
    getOverlayNotifications: lazy('getOverlayNotifications'),
    stats: thrower('stats'),
    mining: thrower('mining'),
    overlayModals: thrower('overlayModals'),
    youtube: thrower('youtube'),
    plugins: { host: thrower('plugins.host') },
    toggleVisibleOverlay: lazy('toggleVisibleOverlay'),
    triggerSubsyncFromConfig: lazy('triggerSubsyncFromConfig'),
    readClipboardText: lazy('readClipboardText'),
    showMpvOsd: lazy('showMpvOsd'),
    sendRendererMpvCommand: lazy('sendRendererMpvCommand'),
    ...overrides,
  };
}

function configWithSubtitleSelection(enabled: boolean): ResolvedConfig {
  return {
    ...DEFAULT_CONFIG,
    subtitleSelection: { ...DEFAULT_CONFIG.subtitleSelection, enabled },
  };
}

test('createSessionActionsAssembly does not call its deps while constructing', () => {
  const sessionActions = createSessionActionsAssembly(makeDeps());
  assert.equal(typeof sessionActions.dispatchSessionAction, 'function');
});

test('the subtitle selection action opens its modal only when subtitle selection is enabled', async () => {
  let enabled = false;
  const opened: string[] = [];
  const sessionActions = createSessionActionsAssembly(
    makeDeps({
      getConfig: () => configWithSubtitleSelection(enabled),
      overlayModals: {
        openOverlayHostedModalWithOsd: (_open: unknown, unavailableMessage: string) => {
          opened.push(unavailableMessage);
        },
      } as unknown as SessionActionsAssemblyDeps['overlayModals'],
    }),
  );

  await sessionActions.dispatchSessionAction({ actionId: 'openSubtitleSelection' });
  assert.deepEqual(opened, []);

  enabled = true;
  await sessionActions.dispatchSessionAction({ actionId: 'openSubtitleSelection' });
  assert.deepEqual(opened, ['Subtitle selection overlay unavailable.']);
});

test('an unregistered plugin action warns and tells the user it is unavailable', async () => {
  const warnings: string[] = [];
  const feedback: string[] = [];
  const sessionActions = createSessionActionsAssembly(
    makeDeps({
      logger: { warn: (message: string) => warnings.push(message) },
      getOverlayNotifications: () =>
        ({
          showConfiguredPlaybackFeedback: (text: string) => feedback.push(text),
        }) as unknown as ReturnType<SessionActionsAssemblyDeps['getOverlayNotifications']>,
      plugins: { host: { runAction: async () => false } },
    }),
  );

  await sessionActions.dispatchSessionAction({
    actionId: 'runPluginAction',
    payload: { pluginActionId: 'example.missing' },
  });

  assert.deepEqual(warnings, ['Plugin action is not registered: example.missing']);
  assert.deepEqual(feedback, ['Plugin action not available: example.missing']);
});

test('renderer mpv commands cycle runtime options only once the options manager exists', () => {
  const sessionActions = createSessionActionsAssembly(makeDeps());

  assert.equal(sessionActions.mpvCommandDeps.hasRuntimeOptionsManager(), false);
  assert.deepEqual(
    sessionActions.mpvCommandDeps.cycleRuntimeOption('subtitle.annotation.jlpt', 1),
    {
      ok: false,
      error: 'Runtime options manager unavailable',
    },
  );
});
