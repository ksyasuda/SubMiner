import type { ResolveContext } from '../context';
import { isObject } from '../shared';
import { applyAiResolution } from './ai';
import { applyLapisResolution } from './lapis';
import { applyModernBehaviorResolution } from './modern-behavior';
import { applyModernFieldsResolution } from './modern-fields';
import { applyModernMediaResolution } from './modern-media';
import { applyModernMetadataResolution } from './modern-metadata';
import { applyProxyResolution } from './proxy';
import { applyTagsResolution } from './tags';

export function applyAnkiModernResolution(
  context: ResolveContext,
  ankiConnect: Record<string, unknown>,
  behavior: Record<string, unknown>,
  media: Record<string, unknown>,
): void {
  const fields = isObject(ankiConnect.fields) ? ankiConnect.fields : {};
  const metadata = isObject(ankiConnect.metadata) ? ankiConnect.metadata : {};

  applyModernFieldsResolution(context, fields);
  applyModernMediaResolution(context, media);
  applyModernBehaviorResolution(context, behavior);
  applyModernMetadataResolution(context, metadata);
  applyLapisResolution(context, ankiConnect);
  applyProxyResolution(context, ankiConnect);
  applyAiResolution(context, ankiConnect);
  applyTagsResolution(context, ankiConnect);
}
