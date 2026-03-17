import { ResolvedConfig } from '../../types';
import { ConfigOptionRegistryEntry } from './shared';

export function buildImmersionConfigOptionRegistry(
  defaultConfig: ResolvedConfig,
): ConfigOptionRegistryEntry[] {
  return [
    {
      path: 'immersionTracking.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.immersionTracking.enabled,
      description: 'Enable immersion tracking for mined subtitle metadata.',
    },
    {
      path: 'immersionTracking.dbPath',
      kind: 'string',
      defaultValue: defaultConfig.immersionTracking.dbPath,
      description:
        'Optional SQLite database path for immersion tracking. Empty value uses the default app data path.',
    },
    {
      path: 'immersionTracking.batchSize',
      kind: 'number',
      defaultValue: defaultConfig.immersionTracking.batchSize,
      description: 'Buffered telemetry/event writes per SQLite transaction.',
    },
    {
      path: 'immersionTracking.flushIntervalMs',
      kind: 'number',
      defaultValue: defaultConfig.immersionTracking.flushIntervalMs,
      description: 'Max delay before queue flush in milliseconds.',
    },
    {
      path: 'immersionTracking.queueCap',
      kind: 'number',
      defaultValue: defaultConfig.immersionTracking.queueCap,
      description: 'In-memory write queue cap before overflow policy applies.',
    },
    {
      path: 'immersionTracking.payloadCapBytes',
      kind: 'number',
      defaultValue: defaultConfig.immersionTracking.payloadCapBytes,
      description: 'Max JSON payload size per event before truncation.',
    },
    {
      path: 'immersionTracking.maintenanceIntervalMs',
      kind: 'number',
      defaultValue: defaultConfig.immersionTracking.maintenanceIntervalMs,
      description: 'Maintenance cadence (prune + rollup + vacuum checks).',
    },
    {
      path: 'immersionTracking.retentionMode',
      kind: 'string',
      defaultValue: defaultConfig.immersionTracking.retentionMode,
      description: 'Retention mode (`preset` uses preset values, `advanced` uses explicit values).',
      enumValues: ['preset', 'advanced'],
    },
    {
      path: 'immersionTracking.retentionPreset',
      kind: 'string',
      defaultValue: defaultConfig.immersionTracking.retentionPreset,
      description: 'Retention preset when `retentionMode` is `preset`.',
      enumValues: ['minimal', 'balanced', 'deep-history'],
    },
    {
      path: 'immersionTracking.retention.eventsDays',
      kind: 'number',
      defaultValue: defaultConfig.immersionTracking.retention.eventsDays,
      description: 'Raw event retention window in days. Use 0 to keep all.',
    },
    {
      path: 'immersionTracking.retention.telemetryDays',
      kind: 'number',
      defaultValue: defaultConfig.immersionTracking.retention.telemetryDays,
      description: 'Telemetry retention window in days. Use 0 to keep all.',
    },
    {
      path: 'immersionTracking.retention.sessionsDays',
      kind: 'number',
      defaultValue: defaultConfig.immersionTracking.retention.sessionsDays,
      description: 'Session retention window in days. Use 0 to keep all.',
    },
    {
      path: 'immersionTracking.retention.dailyRollupsDays',
      kind: 'number',
      defaultValue: defaultConfig.immersionTracking.retention.dailyRollupsDays,
      description: 'Daily rollup retention window in days. Use 0 to keep all.',
    },
    {
      path: 'immersionTracking.retention.monthlyRollupsDays',
      kind: 'number',
      defaultValue: defaultConfig.immersionTracking.retention.monthlyRollupsDays,
      description: 'Monthly rollup retention window in days. Use 0 to keep all.',
    },
    {
      path: 'immersionTracking.retention.vacuumIntervalDays',
      kind: 'number',
      defaultValue: defaultConfig.immersionTracking.retention.vacuumIntervalDays,
      description: 'Minimum days between VACUUM runs. Use 0 to disable.',
    },
    {
      path: 'immersionTracking.lifetimeSummaries.global',
      kind: 'boolean',
      defaultValue: defaultConfig.immersionTracking.lifetimeSummaries?.global,
      description: 'Maintain global lifetime stats rows.',
    },
    {
      path: 'immersionTracking.lifetimeSummaries.anime',
      kind: 'boolean',
      defaultValue: defaultConfig.immersionTracking.lifetimeSummaries?.anime,
      description: 'Maintain per-anime lifetime stats rows.',
    },
    {
      path: 'immersionTracking.lifetimeSummaries.media',
      kind: 'boolean',
      defaultValue: defaultConfig.immersionTracking.lifetimeSummaries?.media,
      description: 'Maintain per-media lifetime stats rows.',
    },
  ];
}
