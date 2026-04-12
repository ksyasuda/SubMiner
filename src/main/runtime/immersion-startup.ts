type ImmersionRetentionPolicy = {
  eventsDays: number;
  telemetryDays: number;
  sessionsDays: number;
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

const DISABLE_IMMERSION_TRACKING_SESSION_ENV = 'SUBMINER_DISABLE_IMMERSION_TRACKING';

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
  const isSessionTrackingDisabled = process.env[DISABLE_IMMERSION_TRACKING_SESSION_ENV] === '1';

  return () => {
    if (isSessionTrackingDisabled) {
      deps.logInfo(
        `Immersion tracking disabled for this session by ${DISABLE_IMMERSION_TRACKING_SESSION_ENV}=1.`,
      );
      return;
    }

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
              sessionsDays: policy.retention.sessionsDays,
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
        try {
          mpvClient.connect();
        } catch (error) {
          deps.logWarn(
            'MPV auto-connect failed during immersion tracker startup; continuing.',
            error,
          );
        }
      }
      deps.seedTrackerFromCurrentMedia();
    } catch (error) {
      deps.logWarn('Immersion tracker startup failed; disabling tracking.', error);
      deps.setTracker(null);
    }
  };
}
