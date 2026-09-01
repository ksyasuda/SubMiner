import type { ResolveContext } from './context';
import { initializeAnkiConnectResolution } from './anki-connect/initialize';
import { applyAnkiKikuResolution } from './anki-connect/kiku';
import { applyAnkiSenrenResolution } from './anki-connect/senren';
import { applyAnkiLapisKikuResolution } from './anki-connect/lapis-kiku';
import { applyAnkiKnownWordsResolution } from './anki-connect/known-words';
import { applyAnkiLegacyResolution } from './anki-connect/legacy';
import { applyAnkiModernResolution } from './anki-connect/modern';
import { isObject } from './shared';

export function applyAnkiConnectResolution(context: ResolveContext): void {
  if (!isObject(context.src.ankiConnect)) {
    return;
  }

  const ankiConnect = context.src.ankiConnect;
  const behavior = isObject(ankiConnect.behavior) ? ankiConnect.behavior : {};
  const fields = isObject(ankiConnect.fields) ? ankiConnect.fields : {};
  const media = isObject(ankiConnect.media) ? ankiConnect.media : {};
  const metadata = isObject(ankiConnect.metadata) ? ankiConnect.metadata : {};

  initializeAnkiConnectResolution(context, ankiConnect);
  applyAnkiModernResolution(context, ankiConnect, behavior, media);
  applyAnkiLegacyResolution(context, ankiConnect, behavior, fields, media, metadata);
  applyAnkiKnownWordsResolution(context, ankiConnect, behavior);
  applyAnkiKikuResolution(context);
  applyAnkiSenrenResolution(context);
  applyAnkiLapisKikuResolution(context, ankiConnect);
}
