import path from 'node:path';

import { mergeAiConfig } from '../ai/config';
import { AnkiIntegration } from '../anki-integration';
import { SubtitleTimingTracker } from '../subtitle-timing-tracker';
import type { ResolvedConfig } from '../types';
import type { AnkiConnectConfig } from '../types/anki';

export async function runHeadlessKnownWordRefresh(input: {
  resolvedConfig: ResolvedConfig;
  runtimeOptionsManager: {
    getEffectiveAnkiConnectConfig: (config: AnkiConnectConfig) => AnkiConnectConfig;
  } | null;
  userDataPath: string;
  logger: {
    error: (message: string, error?: unknown) => void;
  };
  requestAppQuit: () => void;
}): Promise<void> {
  const effectiveAnkiConfig =
    input.runtimeOptionsManager?.getEffectiveAnkiConnectConfig(input.resolvedConfig.ankiConnect) ??
    input.resolvedConfig.ankiConnect;

  if (effectiveAnkiConfig.enabled !== true) {
    input.logger.error('Headless known-word refresh failed: AnkiConnect integration not enabled');
    process.exitCode = 1;
    input.requestAppQuit();
    return;
  }

  const integration = new AnkiIntegration(
    effectiveAnkiConfig,
    new SubtitleTimingTracker(),
    { send: () => undefined } as never,
    undefined,
    undefined,
    async () => ({
      keepNoteId: 0,
      deleteNoteId: 0,
      deleteDuplicate: false,
      cancelled: true,
    }),
    path.join(input.userDataPath, 'known-words-cache.json'),
    mergeAiConfig(input.resolvedConfig.ai, effectiveAnkiConfig.ai),
  );

  try {
    await integration.refreshKnownWordCache();
  } catch (error) {
    input.logger.error('Headless known-word refresh failed:', error);
    process.exitCode = 1;
  } finally {
    integration.stop();
    input.requestAppQuit();
  }
}
