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
    retention: {
      eventsDays: 7,
      telemetryDays: 30,
      dailyRollupsDays: 365,
      monthlyRollupsDays: 5 * 365,
      vacuumIntervalDays: 7,
    },
  },
};
