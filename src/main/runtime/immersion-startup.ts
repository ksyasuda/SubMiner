type ImmersionRetentionPolicy = {
  eventsDays: number;
  telemetryDays: number;
  dailyRollupsDays: number;
  monthlyRollupsDays: number;
  vacuumIntervalDays: number;
};

type ImmersionTrackingPolicy = {
  enabled?: boolean;
  batchSize: number;
  flushIntervalMs: number;
  queueCap: number;
  payloadCapBytes: number;
  maintenanceIntervalMs: number;
  retention: ImmersionRetentionPolicy;
};

type ImmersionTrackingConfig = {
  immersionTracking?: ImmersionTrackingPolicy;
};

type ImmersionTrackerPolicy = Omit<ImmersionTrackingPolicy, 'enabled'>;

type ImmersionTrackerServiceParams = {
  dbPath: string;
  policy: ImmersionTrackerPolicy;
};

type MpvClientLike = {
  connected: boolean;
  connect: () => void;
};

export type ImmersionTrackerStartupDeps = {
  getResolvedConfig: () => ImmersionTrackingConfig;
  getConfiguredDbPath: () => string;
  createTrackerService: (params: ImmersionTrackerServiceParams) => unknown;
  setTracker: (tracker: unknown | null) => void;
  getMpvClient: () => MpvClientLike | null;
  shouldAutoConnectMpv?: () => boolean;
  seedTrackerFromCurrentMedia: () => void;
  logInfo: (message: string) => void;
  logDebug: (message: string) => void;
  logWarn: (message: string, details: unknown) => void;
};

export function createImmersionTrackerStartupHandler(
  deps: ImmersionTrackerStartupDeps,
): () => void {
  return () => {
    const config = deps.getResolvedConfig();
    if (config.immersionTracking?.enabled === false) {
      deps.logInfo('Immersion tracking disabled in config');
      return;
    }

    try {
      deps.logDebug('Immersion tracker startup requested: creating tracker service.');
      const dbPath = deps.getConfiguredDbPath();
      deps.logInfo(`Creating immersion tracker with dbPath=${dbPath}`);

      const policy = config.immersionTracking;
      if (!policy) {
        throw new Error('Immersion tracking policy missing');
      }

      deps.setTracker(
        deps.createTrackerService({
          dbPath,
          policy: {
            batchSize: policy.batchSize,
            flushIntervalMs: policy.flushIntervalMs,
            queueCap: policy.queueCap,
            payloadCapBytes: policy.payloadCapBytes,
            maintenanceIntervalMs: policy.maintenanceIntervalMs,
            retention: {
              eventsDays: policy.retention.eventsDays,
              telemetryDays: policy.retention.telemetryDays,
              dailyRollupsDays: policy.retention.dailyRollupsDays,
              monthlyRollupsDays: policy.retention.monthlyRollupsDays,
              vacuumIntervalDays: policy.retention.vacuumIntervalDays,
            },
          },
        }),
      );
      deps.logDebug('Immersion tracker initialized successfully.');

      const mpvClient = deps.getMpvClient();
      if ((deps.shouldAutoConnectMpv?.() ?? true) && mpvClient && !mpvClient.connected) {
        deps.logInfo('Auto-connecting MPV client for immersion tracking');
        mpvClient.connect();
      }
      deps.seedTrackerFromCurrentMedia();
    } catch (error) {
      deps.logWarn('Immersion tracker startup failed; disabling tracking.', error);
      deps.setTracker(null);
    }
  };
}
