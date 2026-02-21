import { ResolveContext } from './context';
import { asBoolean } from './shared';

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

  if (asBoolean(src.bind_visible_overlay_to_mpv_sub_visibility) !== undefined) {
    resolved.bind_visible_overlay_to_mpv_sub_visibility =
      src.bind_visible_overlay_to_mpv_sub_visibility as boolean;
  } else if (src.bind_visible_overlay_to_mpv_sub_visibility !== undefined) {
    warn(
      'bind_visible_overlay_to_mpv_sub_visibility',
      src.bind_visible_overlay_to_mpv_sub_visibility,
      resolved.bind_visible_overlay_to_mpv_sub_visibility,
      'Expected boolean.',
    );
  }
}
