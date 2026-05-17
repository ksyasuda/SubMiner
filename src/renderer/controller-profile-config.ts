import type { ResolvedControllerConfig, ResolvedControllerProfileConfig } from '../types';

export function getControllerProfile(
  config: ResolvedControllerConfig | null,
  gamepadId: string | null | undefined,
): ResolvedControllerProfileConfig | null {
  if (!config || !gamepadId) return null;
  return config.profiles[gamepadId] ?? null;
}

export function resolveControllerConfigForGamepad(
  config: ResolvedControllerConfig,
  gamepadId: string | null | undefined,
): ResolvedControllerConfig {
  const profile = getControllerProfile(config, gamepadId);
  if (!profile) return config;
  return {
    ...config,
    buttonIndices: profile.buttonIndices,
    bindings: profile.bindings,
  };
}
