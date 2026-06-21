import { ResolveContext } from './context';
import { asBoolean } from './shared';

const VALID_UI_LANGUAGES = new Set(['system', 'en', 'zh-CN']);

export function applyTopLevelConfig(context: ResolveContext): void {
  const { src, resolved, warn } = context;
  const knownTopLevelKeys = new Set(Object.keys(resolved));
  for (const key of Object.keys(src)) {
    if (!knownTopLevelKeys.has(key)) {
      warn(key, src[key], undefined, 'Unknown top-level config key; ignored.');
    }
  }

  if (asBoolean(src.auto_start_overlay) !== undefined) {
    resolved.auto_start_overlay = src.auto_start_overlay as boolean;
  }

  if (src.uiLanguage !== undefined) {
    if (typeof src.uiLanguage === 'string' && VALID_UI_LANGUAGES.has(src.uiLanguage)) {
      resolved.uiLanguage = src.uiLanguage;
    } else {
      warn(
        'uiLanguage',
        src.uiLanguage,
        resolved.uiLanguage,
        "Expected 'system', 'en', or 'zh-CN'.",
      );
    }
  }
}
