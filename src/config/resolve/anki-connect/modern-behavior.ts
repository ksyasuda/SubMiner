import { DEFAULT_CONFIG } from '../../definitions';
import type { ResolveContext } from '../context';
import { asBoolean } from '../shared';
import { applyModernValue } from './modern-value';
import { asNotificationType } from './shared';

export function applyModernBehaviorResolution(
  context: ResolveContext,
  behavior: Record<string, unknown>,
): void {
  for (const key of [
    'overwriteAudio',
    'overwriteImage',
    'highlightWord',
    'autoUpdateNewCards',
  ] as const) {
    applyModernValue(
      context,
      behavior,
      key,
      `ankiConnect.behavior.${key}`,
      asBoolean,
      DEFAULT_CONFIG.ankiConnect.behavior[key],
      (value) => {
        context.resolved.ankiConnect.behavior[key] = value;
      },
      'Expected boolean.',
    );
  }

  applyModernValue(
    context,
    behavior,
    'mediaInsertMode',
    'ankiConnect.behavior.mediaInsertMode',
    (value) => (value === 'append' || value === 'prepend' ? value : undefined),
    DEFAULT_CONFIG.ankiConnect.behavior.mediaInsertMode,
    (value) => {
      context.resolved.ankiConnect.behavior.mediaInsertMode = value;
    },
    "Expected 'append' or 'prepend'.",
  );
  applyModernValue(
    context,
    behavior,
    'notificationType',
    'ankiConnect.behavior.notificationType',
    asNotificationType,
    DEFAULT_CONFIG.ankiConnect.behavior.notificationType,
    (value) => {
      context.resolved.ankiConnect.behavior.notificationType = value;
    },
    "Expected 'overlay', 'system', 'both', 'none', 'osd', or 'osd-system'.",
  );
}
