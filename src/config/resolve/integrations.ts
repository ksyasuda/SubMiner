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

    if (isObject(src.anilist.characterDictionary)) {
      const characterDictionary = src.anilist.characterDictionary;

      const dictionaryEnabled = asBoolean(characterDictionary.enabled);
      if (dictionaryEnabled !== undefined) {
        resolved.anilist.characterDictionary.enabled = dictionaryEnabled;
      } else if (characterDictionary.enabled !== undefined) {
        warn(
          'anilist.characterDictionary.enabled',
          characterDictionary.enabled,
          resolved.anilist.characterDictionary.enabled,
          'Expected boolean.',
        );
      }

      const refreshTtlHours = asNumber(characterDictionary.refreshTtlHours);
      if (refreshTtlHours !== undefined) {
        const normalized = Math.min(24 * 365, Math.max(1, Math.floor(refreshTtlHours)));
        if (normalized !== refreshTtlHours) {
          warn(
            'anilist.characterDictionary.refreshTtlHours',
            characterDictionary.refreshTtlHours,
            normalized,
            'Out of range; clamped to 1..8760 hours.',
          );
        }
        resolved.anilist.characterDictionary.refreshTtlHours = normalized;
      } else if (characterDictionary.refreshTtlHours !== undefined) {
        warn(
          'anilist.characterDictionary.refreshTtlHours',
          characterDictionary.refreshTtlHours,
          resolved.anilist.characterDictionary.refreshTtlHours,
          'Expected number.',
        );
      }

      const maxLoaded = asNumber(characterDictionary.maxLoaded);
      if (maxLoaded !== undefined) {
        const normalized = Math.min(20, Math.max(1, Math.floor(maxLoaded)));
        if (normalized !== maxLoaded) {
          warn(
            'anilist.characterDictionary.maxLoaded',
            characterDictionary.maxLoaded,
            normalized,
            'Out of range; clamped to 1..20.',
          );
        }
        resolved.anilist.characterDictionary.maxLoaded = normalized;
      } else if (characterDictionary.maxLoaded !== undefined) {
        warn(
          'anilist.characterDictionary.maxLoaded',
          characterDictionary.maxLoaded,
          resolved.anilist.characterDictionary.maxLoaded,
          'Expected number.',
        );
      }

      const evictionPolicyRaw = asString(characterDictionary.evictionPolicy);
      if (evictionPolicyRaw !== undefined) {
        const evictionPolicy = evictionPolicyRaw.trim().toLowerCase();
        if (evictionPolicy === 'disable' || evictionPolicy === 'delete') {
          resolved.anilist.characterDictionary.evictionPolicy = evictionPolicy;
        } else {
          warn(
            'anilist.characterDictionary.evictionPolicy',
            characterDictionary.evictionPolicy,
            resolved.anilist.characterDictionary.evictionPolicy,
            "Expected one of: 'disable', 'delete'.",
          );
        }
      } else if (characterDictionary.evictionPolicy !== undefined) {
        warn(
          'anilist.characterDictionary.evictionPolicy',
          characterDictionary.evictionPolicy,
          resolved.anilist.characterDictionary.evictionPolicy,
          'Expected string.',
        );
      }

      const profileScopeRaw = asString(characterDictionary.profileScope);
      if (profileScopeRaw !== undefined) {
        const profileScope = profileScopeRaw.trim().toLowerCase();
        if (profileScope === 'all' || profileScope === 'active') {
          resolved.anilist.characterDictionary.profileScope = profileScope;
        } else {
          warn(
            'anilist.characterDictionary.profileScope',
            characterDictionary.profileScope,
            resolved.anilist.characterDictionary.profileScope,
            "Expected one of: 'all', 'active'.",
          );
        }
      } else if (characterDictionary.profileScope !== undefined) {
        warn(
          'anilist.characterDictionary.profileScope',
          characterDictionary.profileScope,
          resolved.anilist.characterDictionary.profileScope,
          'Expected string.',
        );
      }
    } else if (src.anilist.characterDictionary !== undefined) {
      warn(
        'anilist.characterDictionary',
        src.anilist.characterDictionary,
        resolved.anilist.characterDictionary,
        'Expected object.',
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
