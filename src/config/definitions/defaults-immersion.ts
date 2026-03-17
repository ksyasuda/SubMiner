import { ResolvedConfig } from '../../types';

export const IMMERSION_DEFAULT_CONFIG: Pick<ResolvedConfig, 'immersionTracking'> = {
  immersionTracking: {
    enabled: true,
    dbPath: '',
    batchSize: 25,
    flushIntervalMs: 500,
    queueCap: 1000,
    payloadCapBytes: 256,
    maintenanceIntervalMs: 24 * 60 * 60 * 1000,
    retentionMode: 'preset',
    retentionPreset: 'balanced',
    retention: {
      eventsDays: 0,
      telemetryDays: 0,
      sessionsDays: 0,
      dailyRollupsDays: 0,
      monthlyRollupsDays: 0,
      vacuumIntervalDays: 0,
    },
    lifetimeSummaries: {
      global: true,
      anime: true,
      media: true,
    },
  },
};
