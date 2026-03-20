import { ResolveContext } from './context';
import { ImmersionTrackingRetentionMode, ImmersionTrackingRetentionPreset } from '../../types';
import { asBoolean, asNumber, asString, isObject } from './shared';

const DEFAULT_RETENTION_MODE: ImmersionTrackingRetentionMode = 'preset';
const DEFAULT_RETENTION_PRESET: ImmersionTrackingRetentionPreset = 'balanced';

const BASE_RETENTION = {
  eventsDays: 0,
  telemetryDays: 0,
  sessionsDays: 0,
  dailyRollupsDays: 0,
  monthlyRollupsDays: 0,
  vacuumIntervalDays: 0,
};

const RETENTION_PRESETS: Record<ImmersionTrackingRetentionPreset, typeof BASE_RETENTION> = {
  minimal: {
    eventsDays: 3,
    telemetryDays: 14,
    sessionsDays: 14,
    dailyRollupsDays: 30,
    monthlyRollupsDays: 365,
    vacuumIntervalDays: 7,
  },
  balanced: BASE_RETENTION,
  'deep-history': {
    eventsDays: 14,
    telemetryDays: 60,
    sessionsDays: 60,
    dailyRollupsDays: 730,
    monthlyRollupsDays: 5 * 365,
    vacuumIntervalDays: 7,
  },
};

const DEFAULT_LIFETIME_SUMMARIES = {
  global: true,
  anime: true,
  media: true,
};

function asRetentionMode(value: unknown): value is ImmersionTrackingRetentionMode {
  return value === 'preset' || value === 'advanced';
}

function asRetentionPreset(value: unknown): value is ImmersionTrackingRetentionPreset {
  return value === 'minimal' || value === 'balanced' || value === 'deep-history';
}

export function applyImmersionTrackingConfig(context: ResolveContext): void {
  const { src, resolved, warn } = context;

  if (!isObject(src.immersionTracking)) {
    resolved.immersionTracking.retentionMode = DEFAULT_RETENTION_MODE;
    resolved.immersionTracking.retentionPreset = DEFAULT_RETENTION_PRESET;
    resolved.immersionTracking.retention = {
      ...BASE_RETENTION,
    };
    resolved.immersionTracking.lifetimeSummaries = {
      ...DEFAULT_LIFETIME_SUMMARIES,
    };
    return;
  }

  if (isObject(src.immersionTracking)) {
    const enabled = asBoolean(src.immersionTracking.enabled);
    if (enabled !== undefined) {
      resolved.immersionTracking.enabled = enabled;
    } else if (src.immersionTracking.enabled !== undefined) {
      warn(
        'immersionTracking.enabled',
        src.immersionTracking.enabled,
        resolved.immersionTracking.enabled,
        'Expected boolean.',
      );
    }

    const dbPath = asString(src.immersionTracking.dbPath);
    if (dbPath !== undefined) {
      resolved.immersionTracking.dbPath = dbPath;
    } else if (src.immersionTracking.dbPath !== undefined) {
      warn(
        'immersionTracking.dbPath',
        src.immersionTracking.dbPath,
        resolved.immersionTracking.dbPath,
        'Expected string.',
      );
    }

    const batchSize = asNumber(src.immersionTracking.batchSize);
    if (batchSize !== undefined && batchSize >= 1 && batchSize <= 10_000) {
      resolved.immersionTracking.batchSize = Math.floor(batchSize);
    } else if (src.immersionTracking.batchSize !== undefined) {
      warn(
        'immersionTracking.batchSize',
        src.immersionTracking.batchSize,
        resolved.immersionTracking.batchSize,
        'Expected integer between 1 and 10000.',
      );
    }

    const flushIntervalMs = asNumber(src.immersionTracking.flushIntervalMs);
    if (flushIntervalMs !== undefined && flushIntervalMs >= 50 && flushIntervalMs <= 60_000) {
      resolved.immersionTracking.flushIntervalMs = Math.floor(flushIntervalMs);
    } else if (src.immersionTracking.flushIntervalMs !== undefined) {
      warn(
        'immersionTracking.flushIntervalMs',
        src.immersionTracking.flushIntervalMs,
        resolved.immersionTracking.flushIntervalMs,
        'Expected integer between 50 and 60000.',
      );
    }

    const queueCap = asNumber(src.immersionTracking.queueCap);
    if (queueCap !== undefined && queueCap >= 100 && queueCap <= 100_000) {
      resolved.immersionTracking.queueCap = Math.floor(queueCap);
    } else if (src.immersionTracking.queueCap !== undefined) {
      warn(
        'immersionTracking.queueCap',
        src.immersionTracking.queueCap,
        resolved.immersionTracking.queueCap,
        'Expected integer between 100 and 100000.',
      );
    }

    const payloadCapBytes = asNumber(src.immersionTracking.payloadCapBytes);
    if (payloadCapBytes !== undefined && payloadCapBytes >= 64 && payloadCapBytes <= 8192) {
      resolved.immersionTracking.payloadCapBytes = Math.floor(payloadCapBytes);
    } else if (src.immersionTracking.payloadCapBytes !== undefined) {
      warn(
        'immersionTracking.payloadCapBytes',
        src.immersionTracking.payloadCapBytes,
        resolved.immersionTracking.payloadCapBytes,
        'Expected integer between 64 and 8192.',
      );
    }

    const maintenanceIntervalMs = asNumber(src.immersionTracking.maintenanceIntervalMs);
    if (
      maintenanceIntervalMs !== undefined &&
      maintenanceIntervalMs >= 60_000 &&
      maintenanceIntervalMs <= 7 * 24 * 60 * 60 * 1000
    ) {
      resolved.immersionTracking.maintenanceIntervalMs = Math.floor(maintenanceIntervalMs);
    } else if (src.immersionTracking.maintenanceIntervalMs !== undefined) {
      warn(
        'immersionTracking.maintenanceIntervalMs',
        src.immersionTracking.maintenanceIntervalMs,
        resolved.immersionTracking.maintenanceIntervalMs,
        'Expected integer between 60000 and 604800000.',
      );
    }

    const retentionMode = asString(src.immersionTracking.retentionMode);
    if (asRetentionMode(retentionMode)) {
      resolved.immersionTracking.retentionMode = retentionMode;
    } else if (src.immersionTracking.retentionMode !== undefined) {
      warn(
        'immersionTracking.retentionMode',
        src.immersionTracking.retentionMode,
        DEFAULT_RETENTION_MODE,
        'Expected "preset" or "advanced".',
      );
      resolved.immersionTracking.retentionMode = DEFAULT_RETENTION_MODE;
    } else {
      resolved.immersionTracking.retentionMode = DEFAULT_RETENTION_MODE;
    }

    const retentionPreset = asString(src.immersionTracking.retentionPreset);
    if (asRetentionPreset(retentionPreset)) {
      resolved.immersionTracking.retentionPreset = retentionPreset;
    } else if (src.immersionTracking.retentionPreset !== undefined) {
      warn(
        'immersionTracking.retentionPreset',
        src.immersionTracking.retentionPreset,
        DEFAULT_RETENTION_PRESET,
        'Expected "minimal", "balanced", or "deep-history".',
      );
      resolved.immersionTracking.retentionPreset = DEFAULT_RETENTION_PRESET;
    } else {
      resolved.immersionTracking.retentionPreset =
        resolved.immersionTracking.retentionPreset ?? DEFAULT_RETENTION_PRESET;
    }

    const resolvedPreset =
      resolved.immersionTracking.retentionPreset === 'minimal' ||
      resolved.immersionTracking.retentionPreset === 'balanced' ||
      resolved.immersionTracking.retentionPreset === 'deep-history'
        ? resolved.immersionTracking.retentionPreset
        : DEFAULT_RETENTION_PRESET;

    const baseRetention =
      resolved.immersionTracking.retentionMode === 'preset'
        ? RETENTION_PRESETS[resolvedPreset]
        : BASE_RETENTION;

    const retention = {
      eventsDays: baseRetention.eventsDays,
      telemetryDays: baseRetention.telemetryDays,
      sessionsDays: baseRetention.sessionsDays,
      dailyRollupsDays: baseRetention.dailyRollupsDays,
      monthlyRollupsDays: baseRetention.monthlyRollupsDays,
      vacuumIntervalDays: baseRetention.vacuumIntervalDays,
    };

    if (isObject(src.immersionTracking.retention)) {
      const eventsDays = asNumber(src.immersionTracking.retention.eventsDays);
      if (eventsDays !== undefined && eventsDays >= 0 && eventsDays <= 3650) {
        retention.eventsDays = Math.floor(eventsDays);
      } else if (src.immersionTracking.retention.eventsDays !== undefined) {
        warn(
          'immersionTracking.retention.eventsDays',
          src.immersionTracking.retention.eventsDays,
          retention.eventsDays,
          'Expected integer between 0 and 3650.',
        );
      }

      const telemetryDays = asNumber(src.immersionTracking.retention.telemetryDays);
      if (telemetryDays !== undefined && telemetryDays >= 0 && telemetryDays <= 3650) {
        retention.telemetryDays = Math.floor(telemetryDays);
      } else if (src.immersionTracking.retention.telemetryDays !== undefined) {
        warn(
          'immersionTracking.retention.telemetryDays',
          src.immersionTracking.retention.telemetryDays,
          retention.telemetryDays,
          'Expected integer between 0 and 3650.',
        );
      }

      const sessionsDays = asNumber(src.immersionTracking.retention.sessionsDays);
      if (sessionsDays !== undefined && sessionsDays >= 0 && sessionsDays <= 3650) {
        retention.sessionsDays = Math.floor(sessionsDays);
      } else if (src.immersionTracking.retention.sessionsDays !== undefined) {
        warn(
          'immersionTracking.retention.sessionsDays',
          src.immersionTracking.retention.sessionsDays,
          retention.sessionsDays,
          'Expected integer between 0 and 3650.',
        );
      }

      const dailyRollupsDays = asNumber(src.immersionTracking.retention.dailyRollupsDays);
      if (dailyRollupsDays !== undefined && dailyRollupsDays >= 0 && dailyRollupsDays <= 36500) {
        retention.dailyRollupsDays = Math.floor(dailyRollupsDays);
      } else if (src.immersionTracking.retention.dailyRollupsDays !== undefined) {
        warn(
          'immersionTracking.retention.dailyRollupsDays',
          src.immersionTracking.retention.dailyRollupsDays,
          retention.dailyRollupsDays,
          'Expected integer between 0 and 36500.',
        );
      }

      const monthlyRollupsDays = asNumber(src.immersionTracking.retention.monthlyRollupsDays);
      if (
        monthlyRollupsDays !== undefined &&
        monthlyRollupsDays >= 0 &&
        monthlyRollupsDays <= 36500
      ) {
        retention.monthlyRollupsDays = Math.floor(monthlyRollupsDays);
      } else if (src.immersionTracking.retention.monthlyRollupsDays !== undefined) {
        warn(
          'immersionTracking.retention.monthlyRollupsDays',
          src.immersionTracking.retention.monthlyRollupsDays,
          retention.monthlyRollupsDays,
          'Expected integer between 0 and 36500.',
        );
      }

      const vacuumIntervalDays = asNumber(src.immersionTracking.retention.vacuumIntervalDays);
      if (
        vacuumIntervalDays !== undefined &&
        vacuumIntervalDays >= 0 &&
        vacuumIntervalDays <= 3650
      ) {
        retention.vacuumIntervalDays = Math.floor(vacuumIntervalDays);
      } else if (src.immersionTracking.retention.vacuumIntervalDays !== undefined) {
        warn(
          'immersionTracking.retention.vacuumIntervalDays',
          src.immersionTracking.retention.vacuumIntervalDays,
          retention.vacuumIntervalDays,
          'Expected integer between 0 and 3650.',
        );
      }
    } else if (src.immersionTracking.retention !== undefined) {
      warn(
        'immersionTracking.retention',
        src.immersionTracking.retention,
        baseRetention,
        'Expected object.',
      );
    }

    resolved.immersionTracking.retention = {
      eventsDays: retention.eventsDays,
      telemetryDays: retention.telemetryDays,
      sessionsDays: retention.sessionsDays,
      dailyRollupsDays: retention.dailyRollupsDays,
      monthlyRollupsDays: retention.monthlyRollupsDays,
      vacuumIntervalDays: retention.vacuumIntervalDays,
    };

    const lifetimeSummaries = {
      global: DEFAULT_LIFETIME_SUMMARIES.global,
      anime: DEFAULT_LIFETIME_SUMMARIES.anime,
      media: DEFAULT_LIFETIME_SUMMARIES.media,
    };

    if (isObject(src.immersionTracking.lifetimeSummaries)) {
      const global = asBoolean(src.immersionTracking.lifetimeSummaries.global);
      if (global !== undefined) {
        lifetimeSummaries.global = global;
      }

      const anime = asBoolean(src.immersionTracking.lifetimeSummaries.anime);
      if (anime !== undefined) {
        lifetimeSummaries.anime = anime;
      }

      const media = asBoolean(src.immersionTracking.lifetimeSummaries.media);
      if (media !== undefined) {
        lifetimeSummaries.media = media;
      }
    } else if (src.immersionTracking.lifetimeSummaries !== undefined) {
      warn(
        'immersionTracking.lifetimeSummaries',
        src.immersionTracking.lifetimeSummaries,
        DEFAULT_LIFETIME_SUMMARIES,
        'Expected object.',
      );
    }

    resolved.immersionTracking.lifetimeSummaries = lifetimeSummaries;
  }
}
