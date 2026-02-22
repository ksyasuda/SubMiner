import { ResolveContext } from './context';
import { asBoolean, asNumber, asString, isObject } from './shared';

export function applyIntegrationConfig(context: ResolveContext): void {
  const { src, resolved, warn } = context;

  if (isObject(src.anilist)) {
    const enabled = asBoolean(src.anilist.enabled);
    if (enabled !== undefined) {
      resolved.anilist.enabled = enabled;
    } else if (src.anilist.enabled !== undefined) {
      warn('anilist.enabled', src.anilist.enabled, resolved.anilist.enabled, 'Expected boolean.');
    }

    const accessToken = asString(src.anilist.accessToken);
    if (accessToken !== undefined) {
      resolved.anilist.accessToken = accessToken;
    } else if (src.anilist.accessToken !== undefined) {
      warn(
        'anilist.accessToken',
        src.anilist.accessToken,
        resolved.anilist.accessToken,
        'Expected string.',
      );
    }
  }

  if (isObject(src.jellyfin)) {
    const enabled = asBoolean(src.jellyfin.enabled);
    if (enabled !== undefined) {
      resolved.jellyfin.enabled = enabled;
    } else if (src.jellyfin.enabled !== undefined) {
      warn(
        'jellyfin.enabled',
        src.jellyfin.enabled,
        resolved.jellyfin.enabled,
        'Expected boolean.',
      );
    }

    const stringKeys = [
      'serverUrl',
      'username',
      'deviceId',
      'clientName',
      'clientVersion',
      'defaultLibraryId',
      'iconCacheDir',
      'transcodeVideoCodec',
    ] as const;
    for (const key of stringKeys) {
      const value = asString(src.jellyfin[key]);
      if (value !== undefined) {
        resolved.jellyfin[key] = value as (typeof resolved.jellyfin)[typeof key];
      } else if (src.jellyfin[key] !== undefined) {
        warn(`jellyfin.${key}`, src.jellyfin[key], resolved.jellyfin[key], 'Expected string.');
      }
    }

    const booleanKeys = [
      'remoteControlEnabled',
      'remoteControlAutoConnect',
      'autoAnnounce',
      'directPlayPreferred',
      'pullPictures',
    ] as const;
    for (const key of booleanKeys) {
      const value = asBoolean(src.jellyfin[key]);
      if (value !== undefined) {
        resolved.jellyfin[key] = value as (typeof resolved.jellyfin)[typeof key];
      } else if (src.jellyfin[key] !== undefined) {
        warn(`jellyfin.${key}`, src.jellyfin[key], resolved.jellyfin[key], 'Expected boolean.');
      }
    }

    if (Array.isArray(src.jellyfin.directPlayContainers)) {
      resolved.jellyfin.directPlayContainers = src.jellyfin.directPlayContainers
        .filter((item): item is string => typeof item === 'string')
        .map((item) => item.trim().toLowerCase())
        .filter((item) => item.length > 0);
    } else if (src.jellyfin.directPlayContainers !== undefined) {
      warn(
        'jellyfin.directPlayContainers',
        src.jellyfin.directPlayContainers,
        resolved.jellyfin.directPlayContainers,
        'Expected string array.',
      );
    }
  }

  if (isObject(src.discordPresence)) {
    const enabled = asBoolean(src.discordPresence.enabled);
    if (enabled !== undefined) {
      resolved.discordPresence.enabled = enabled;
    } else if (src.discordPresence.enabled !== undefined) {
      warn(
        'discordPresence.enabled',
        src.discordPresence.enabled,
        resolved.discordPresence.enabled,
        'Expected boolean.',
      );
    }

    const stringKeys = [
      'clientId',
      'detailsTemplate',
      'stateTemplate',
      'largeImageKey',
      'largeImageText',
      'smallImageKey',
      'smallImageText',
      'buttonLabel',
      'buttonUrl',
    ] as const;
    for (const key of stringKeys) {
      const value = asString(src.discordPresence[key]);
      if (value !== undefined) {
        resolved.discordPresence[key] = value;
      } else if (src.discordPresence[key] !== undefined) {
        warn(
          `discordPresence.${key}`,
          src.discordPresence[key],
          resolved.discordPresence[key],
          'Expected string.',
        );
      }
    }

    const updateIntervalMs = asNumber(src.discordPresence.updateIntervalMs);
    if (updateIntervalMs !== undefined) {
      resolved.discordPresence.updateIntervalMs = Math.max(1_000, Math.floor(updateIntervalMs));
    } else if (src.discordPresence.updateIntervalMs !== undefined) {
      warn(
        'discordPresence.updateIntervalMs',
        src.discordPresence.updateIntervalMs,
        resolved.discordPresence.updateIntervalMs,
        'Expected number.',
      );
    }

    const debounceMs = asNumber(src.discordPresence.debounceMs);
    if (debounceMs !== undefined) {
      resolved.discordPresence.debounceMs = Math.max(0, Math.floor(debounceMs));
    } else if (src.discordPresence.debounceMs !== undefined) {
      warn(
        'discordPresence.debounceMs',
        src.discordPresence.debounceMs,
        resolved.discordPresence.debounceMs,
        'Expected number.',
      );
    }
  }
}
