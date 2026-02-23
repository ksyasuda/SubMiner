import { ResolveContext } from './context';
import { asBoolean, asNumber, asString, isObject } from './shared';

export function applyImmersionTrackingConfig(context: ResolveContext): void {
  const { src, resolved, warn } = context;

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

    if (isObject(src.immersionTracking.retention)) {
      const eventsDays = asNumber(src.immersionTracking.retention.eventsDays);
      if (eventsDays !== undefined && eventsDays >= 1 && eventsDays <= 3650) {
        resolved.immersionTracking.retention.eventsDays = Math.floor(eventsDays);
      } else if (src.immersionTracking.retention.eventsDays !== undefined) {
        warn(
          'immersionTracking.retention.eventsDays',
          src.immersionTracking.retention.eventsDays,
          resolved.immersionTracking.retention.eventsDays,
          'Expected integer between 1 and 3650.',
        );
      }

      const telemetryDays = asNumber(src.immersionTracking.retention.telemetryDays);
      if (telemetryDays !== undefined && telemetryDays >= 1 && telemetryDays <= 3650) {
        resolved.immersionTracking.retention.telemetryDays = Math.floor(telemetryDays);
      } else if (src.immersionTracking.retention.telemetryDays !== undefined) {
        warn(
          'immersionTracking.retention.telemetryDays',
          src.immersionTracking.retention.telemetryDays,
          resolved.immersionTracking.retention.telemetryDays,
          'Expected integer between 1 and 3650.',
        );
      }

      const dailyRollupsDays = asNumber(src.immersionTracking.retention.dailyRollupsDays);
      if (dailyRollupsDays !== undefined && dailyRollupsDays >= 1 && dailyRollupsDays <= 36500) {
        resolved.immersionTracking.retention.dailyRollupsDays = Math.floor(dailyRollupsDays);
      } else if (src.immersionTracking.retention.dailyRollupsDays !== undefined) {
        warn(
          'immersionTracking.retention.dailyRollupsDays',
          src.immersionTracking.retention.dailyRollupsDays,
          resolved.immersionTracking.retention.dailyRollupsDays,
          'Expected integer between 1 and 36500.',
        );
      }

      const monthlyRollupsDays = asNumber(src.immersionTracking.retention.monthlyRollupsDays);
      if (
        monthlyRollupsDays !== undefined &&
        monthlyRollupsDays >= 1 &&
        monthlyRollupsDays <= 36500
      ) {
        resolved.immersionTracking.retention.monthlyRollupsDays = Math.floor(monthlyRollupsDays);
      } else if (src.immersionTracking.retention.monthlyRollupsDays !== undefined) {
        warn(
          'immersionTracking.retention.monthlyRollupsDays',
          src.immersionTracking.retention.monthlyRollupsDays,
          resolved.immersionTracking.retention.monthlyRollupsDays,
          'Expected integer between 1 and 36500.',
        );
      }

      const vacuumIntervalDays = asNumber(src.immersionTracking.retention.vacuumIntervalDays);
      if (
        vacuumIntervalDays !== undefined &&
        vacuumIntervalDays >= 1 &&
        vacuumIntervalDays <= 3650
      ) {
        resolved.immersionTracking.retention.vacuumIntervalDays = Math.floor(vacuumIntervalDays);
      } else if (src.immersionTracking.retention.vacuumIntervalDays !== undefined) {
        warn(
          'immersionTracking.retention.vacuumIntervalDays',
          src.immersionTracking.retention.vacuumIntervalDays,
          resolved.immersionTracking.retention.vacuumIntervalDays,
          'Expected integer between 1 and 3650.',
        );
      }
    } else if (src.immersionTracking.retention !== undefined) {
      warn(
        'immersionTracking.retention',
        src.immersionTracking.retention,
        resolved.immersionTracking.retention,
        'Expected object.',
      );
    }
  }
}
