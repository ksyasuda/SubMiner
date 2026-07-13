import { ResolveContext } from './context';
import { asBoolean } from './shared';

export function applyTopLevelConfig(context: ResolveContext): void {
  const { src, resolved, warn } = context;
  const knownTopLevelKeys = new Set([...Object.keys(resolved), 'animetosho']);
  for (const key of Object.keys(src)) {
    if (!knownTopLevelKeys.has(key)) {
      warn(key, src[key], undefined, 'Unknown top-level config key; ignored.');
    }
  }

  if (asBoolean(src.auto_start_overlay) !== undefined) {
    resolved.auto_start_overlay = src.auto_start_overlay as boolean;
  }
}
