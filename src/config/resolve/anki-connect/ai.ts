import type { ResolveContext } from '../context';
import { asBoolean, asString, isObject } from '../shared';

export function applyAiResolution(
  context: ResolveContext,
  ankiConnect: Record<string, unknown>,
): void {
  if (isObject(ankiConnect.ai)) {
    const aiEnabled = asBoolean(ankiConnect.ai.enabled);
    if (aiEnabled !== undefined) {
      context.resolved.ankiConnect.ai.enabled = aiEnabled;
    } else if (ankiConnect.ai.enabled !== undefined) {
      context.warn(
        'ankiConnect.ai.enabled',
        ankiConnect.ai.enabled,
        context.resolved.ankiConnect.ai.enabled,
        'Expected boolean.',
      );
    }

    const aiModel = asString(ankiConnect.ai.model);
    if (aiModel !== undefined) {
      context.resolved.ankiConnect.ai.model = aiModel;
    } else if (ankiConnect.ai.model !== undefined) {
      context.warn(
        'ankiConnect.ai.model',
        ankiConnect.ai.model,
        context.resolved.ankiConnect.ai.model,
        'Expected string.',
      );
    }

    const aiSystemPrompt = asString(ankiConnect.ai.systemPrompt);
    if (aiSystemPrompt !== undefined) {
      context.resolved.ankiConnect.ai.systemPrompt = aiSystemPrompt;
    } else if (ankiConnect.ai.systemPrompt !== undefined) {
      context.warn(
        'ankiConnect.ai.systemPrompt',
        ankiConnect.ai.systemPrompt,
        context.resolved.ankiConnect.ai.systemPrompt,
        'Expected string.',
      );
    }
  } else {
    const aiEnabled = asBoolean(ankiConnect.ai);
    if (aiEnabled !== undefined) {
      context.resolved.ankiConnect.ai.enabled = aiEnabled;
    } else if (ankiConnect.ai !== undefined) {
      context.warn(
        'ankiConnect.ai',
        ankiConnect.ai,
        context.resolved.ankiConnect.ai.enabled,
        'Expected boolean or object.',
      );
    }
  }
}
