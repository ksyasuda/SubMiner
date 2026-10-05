import { parseArgs } from '../config.js';
import type { Args } from '../types.js';

/** Default launcher `Args` (as if run with no CLI input), with `overrides` applied. */
export function makeLauncherArgs(overrides: Partial<Args> = {}): Args {
  return { ...parseArgs([], 'subminer', {}), ...overrides };
}
