import { DEFAULT_CONFIG } from '../../definitions';
import type { ResolveContext } from '../context';

export function applyTagsResolution(
  context: ResolveContext,
  ankiConnect: Record<string, unknown>,
): void {
  if (Array.isArray(ankiConnect.tags)) {
    const normalizedTags = ankiConnect.tags
      .filter((entry): entry is string => typeof entry === 'string')
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0);
    if (normalizedTags.length === ankiConnect.tags.length) {
      context.resolved.ankiConnect.tags = [...new Set(normalizedTags)];
    } else {
      context.resolved.ankiConnect.tags = DEFAULT_CONFIG.ankiConnect.tags;
      context.warn(
        'ankiConnect.tags',
        ankiConnect.tags,
        context.resolved.ankiConnect.tags,
        'Expected an array of non-empty strings.',
      );
    }
  } else if (ankiConnect.tags !== undefined) {
    context.resolved.ankiConnect.tags = DEFAULT_CONFIG.ankiConnect.tags;
    context.warn(
      'ankiConnect.tags',
      ankiConnect.tags,
      context.resolved.ankiConnect.tags,
      'Expected an array of strings.',
    );
  }
}
