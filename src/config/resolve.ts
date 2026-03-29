import { ConfigValidationWarning, RawConfig, ResolvedConfig } from '../types/config';
import { applyAnkiConnectResolution } from './resolve/anki-connect';
import { createResolveContext } from './resolve/context';
import { applyCoreDomainConfig } from './resolve/core-domains';
import { applyImmersionTrackingConfig } from './resolve/immersion-tracking';
import { applyIntegrationConfig } from './resolve/integrations';
import { applyStatsConfig } from './resolve/stats';
import { applySubtitleDomainConfig } from './resolve/subtitle-domains';
import { applyTopLevelConfig } from './resolve/top-level';

const APPLY_RESOLVE_STEPS = [
  applyTopLevelConfig,
  applyCoreDomainConfig,
  applySubtitleDomainConfig,
  applyIntegrationConfig,
  applyImmersionTrackingConfig,
  applyStatsConfig,
  applyAnkiConnectResolution,
] as const;

export function resolveConfig(raw: RawConfig): {
  resolved: ResolvedConfig;
  warnings: ConfigValidationWarning[];
} {
  const { context, warnings } = createResolveContext(raw);

  for (const applyStep of APPLY_RESOLVE_STEPS) {
    applyStep(context);
  }

  return {
    resolved: context.resolved,
    warnings,
  };
}
