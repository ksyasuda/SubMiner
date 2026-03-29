import { ConfigValidationWarning, RawConfig, ResolvedConfig } from '../../types/config';
import { DEFAULT_CONFIG, deepCloneConfig } from '../definitions';
import { createWarningCollector } from '../warnings';
import { isObject } from './shared';

export interface ResolveContext {
  src: Record<string, unknown>;
  resolved: ResolvedConfig;
  warn(path: string, value: unknown, fallback: unknown, message: string): void;
}

export type ResolveConfigApplier = (context: ResolveContext) => void;

export function createResolveContext(raw: RawConfig): {
  context: ResolveContext;
  warnings: ConfigValidationWarning[];
} {
  const resolved = deepCloneConfig(DEFAULT_CONFIG);
  const { warnings, warn } = createWarningCollector();
  const src = isObject(raw) ? raw : {};

  return {
    context: {
      src,
      resolved,
      warn,
    },
    warnings,
  };
}
