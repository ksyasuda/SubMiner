import * as os from 'node:os';
import * as path from 'node:path';
import { MPV_LAUNCH_MODE_VALUES, parseMpvLaunchMode } from '../../shared/mpv-launch-mode';
import { ResolveContext } from './context';
import { asBoolean, asNumber, asString, isObject } from './shared';
import { isValidRepoUrl } from '../../anime-bridge/extension-repo';

function normalizeExternalProfilePath(value: string): string {
  const trimmed = value.trim();
  if (trimmed === '~') {
    return os.homedir();
  }
  if (trimmed.startsWith('~/') || trimmed.startsWith('~\\')) {
    return path.join(os.homedir(), trimmed.slice(2));
  }
  return trimmed;
}

export function applyIntegrationConfig(context: ResolveContext): void {
  const { src, resolved, warn } = context;

  if (isObject(src.ai)) {
    const booleanKeys = ['enabled'] as const;
    for (const key of booleanKeys) {
      const value = asBoolean(src.ai[key]);
      if (value !== undefined) {
        resolved.ai[key] = value;
      } else if (src.ai[key] !== undefined) {
        warn(`ai.${key}`, src.ai[key], resolved.ai[key], 'Expected boolean.');
      }
    }

    const stringKeys = ['apiKey', 'apiKeyCommand', 'baseUrl', 'model', 'systemPrompt'] as const;
    for (const key of stringKeys) {
      const value = asString(src.ai[key]);
      if (value !== undefined) {
        resolved.ai[key] = value;
      } else if (src.ai[key] !== undefined) {
        warn(`ai.${key}`, src.ai[key], resolved.ai[key], 'Expected string.');
      }
    }

    const requestTimeoutMs = asNumber(src.ai.requestTimeoutMs);
    if (
      requestTimeoutMs !== undefined &&
      Number.isInteger(requestTimeoutMs) &&
      requestTimeoutMs > 0
    ) {
      resolved.ai.requestTimeoutMs = requestTimeoutMs;
    } else if (src.ai.requestTimeoutMs !== undefined) {
      warn(
        'ai.requestTimeoutMs',
        src.ai.requestTimeoutMs,
        resolved.ai.requestTimeoutMs,
        'Expected positive integer.',
      );
    }
  } else if (src.ai !== undefined) {
    warn('ai', src.ai, resolved.ai, 'Expected object.');
  }

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

      if (isObject(characterDictionary.collapsibleSections)) {
        const collapsibleSections = characterDictionary.collapsibleSections;
        const keys = ['description', 'characterInformation', 'voicedBy'] as const;
        for (const key of keys) {
          const value = asBoolean(collapsibleSections[key]);
          if (value !== undefined) {
            resolved.anilist.characterDictionary.collapsibleSections[key] = value;
          } else if (collapsibleSections[key] !== undefined) {
            warn(
              `anilist.characterDictionary.collapsibleSections.${key}`,
              collapsibleSections[key],
              resolved.anilist.characterDictionary.collapsibleSections[key],
              'Expected boolean.',
            );
          }
        }
      } else if (characterDictionary.collapsibleSections !== undefined) {
        warn(
          'anilist.characterDictionary.collapsibleSections',
          characterDictionary.collapsibleSections,
          resolved.anilist.characterDictionary.collapsibleSections,
          'Expected object.',
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

  if (isObject(src.yomitan)) {
    const externalProfilePath = asString(src.yomitan.externalProfilePath);
    if (externalProfilePath !== undefined) {
      resolved.yomitan.externalProfilePath = normalizeExternalProfilePath(externalProfilePath);
    } else if (src.yomitan.externalProfilePath !== undefined) {
      warn(
        'yomitan.externalProfilePath',
        src.yomitan.externalProfilePath,
        resolved.yomitan.externalProfilePath,
        'Expected string.',
      );
    }
  } else if (src.yomitan !== undefined) {
    warn('yomitan', src.yomitan, resolved.yomitan, 'Expected object.');
  }

  if (isObject(src.mpv)) {
    const executablePath = asString(src.mpv.executablePath);
    if (executablePath !== undefined) {
      resolved.mpv.executablePath = executablePath.trim();
    } else if (src.mpv.executablePath !== undefined) {
      warn(
        'mpv.executablePath',
        src.mpv.executablePath,
        resolved.mpv.executablePath,
        'Expected string.',
      );
    }

    const launchMode = parseMpvLaunchMode(src.mpv.launchMode);
    if (launchMode !== undefined) {
      resolved.mpv.launchMode = launchMode;
    } else if (src.mpv.launchMode !== undefined) {
      warn(
        'mpv.launchMode',
        src.mpv.launchMode,
        resolved.mpv.launchMode,
        `Expected one of: ${MPV_LAUNCH_MODE_VALUES.map((value) => `'${value}'`).join(', ')}.`,
      );
    }

    const profile = asString(src.mpv.profile);
    if (profile !== undefined) {
      resolved.mpv.profile = profile.trim();
    } else if (src.mpv.profile !== undefined) {
      warn('mpv.profile', src.mpv.profile, resolved.mpv.profile, 'Expected string.');
    }

    const socketPath = asString(src.mpv.socketPath);
    if (socketPath !== undefined && socketPath.trim().length > 0) {
      resolved.mpv.socketPath = socketPath.trim();
    } else if (src.mpv.socketPath !== undefined) {
      warn(
        'mpv.socketPath',
        src.mpv.socketPath,
        resolved.mpv.socketPath,
        'Expected non-empty string.',
      );
    }

    const backend = asString(src.mpv.backend);
    if (
      backend === 'auto' ||
      backend === 'hyprland' ||
      backend === 'sway' ||
      backend === 'x11' ||
      backend === 'macos' ||
      backend === 'windows'
    ) {
      resolved.mpv.backend = backend;
    } else if (src.mpv.backend !== undefined) {
      warn(
        'mpv.backend',
        src.mpv.backend,
        resolved.mpv.backend,
        'Expected auto, hyprland, sway, x11, macos, or windows.',
      );
    }

    const autoStartSubMiner = asBoolean(src.mpv.autoStartSubMiner);
    if (autoStartSubMiner !== undefined) {
      resolved.mpv.autoStartSubMiner = autoStartSubMiner;
    } else if (src.mpv.autoStartSubMiner !== undefined) {
      warn(
        'mpv.autoStartSubMiner',
        src.mpv.autoStartSubMiner,
        resolved.mpv.autoStartSubMiner,
        'Expected boolean.',
      );
    }

    const pauseUntilOverlayReady = asBoolean(src.mpv.pauseUntilOverlayReady);
    if (pauseUntilOverlayReady !== undefined) {
      resolved.mpv.pauseUntilOverlayReady = pauseUntilOverlayReady;
    } else if (src.mpv.pauseUntilOverlayReady !== undefined) {
      warn(
        'mpv.pauseUntilOverlayReady',
        src.mpv.pauseUntilOverlayReady,
        resolved.mpv.pauseUntilOverlayReady,
        'Expected boolean.',
      );
    }

    const subminerBinaryPath = asString(src.mpv.subminerBinaryPath);
    if (subminerBinaryPath !== undefined) {
      resolved.mpv.subminerBinaryPath = subminerBinaryPath.trim();
    } else if (src.mpv.subminerBinaryPath !== undefined) {
      warn(
        'mpv.subminerBinaryPath',
        src.mpv.subminerBinaryPath,
        resolved.mpv.subminerBinaryPath,
        'Expected string.',
      );
    }

    const aniskipEnabled = asBoolean(src.mpv.aniskipEnabled);
    if (aniskipEnabled !== undefined) {
      resolved.mpv.aniskipEnabled = aniskipEnabled;
    } else if (src.mpv.aniskipEnabled !== undefined) {
      warn(
        'mpv.aniskipEnabled',
        src.mpv.aniskipEnabled,
        resolved.mpv.aniskipEnabled,
        'Expected boolean.',
      );
    }

    const aniskipButtonKey = asString(src.mpv.aniskipButtonKey);
    if (aniskipButtonKey !== undefined && aniskipButtonKey.trim().length > 0) {
      resolved.mpv.aniskipButtonKey = aniskipButtonKey.trim();
    } else if (src.mpv.aniskipButtonKey !== undefined) {
      warn(
        'mpv.aniskipButtonKey',
        src.mpv.aniskipButtonKey,
        resolved.mpv.aniskipButtonKey,
        'Expected non-empty string.',
      );
    }
  } else if (src.mpv !== undefined) {
    warn('mpv', src.mpv, resolved.mpv, 'Expected object.');
  }

  if (isObject(src.anime)) {
    const extensionsDir = asString(src.anime.extensionsDir);
    if (extensionsDir !== undefined) {
      resolved.anime.extensionsDir = normalizeExternalProfilePath(extensionsDir);
    } else if (src.anime.extensionsDir !== undefined) {
      warn(
        'anime.extensionsDir',
        src.anime.extensionsDir,
        resolved.anime.extensionsDir,
        'Expected string.',
      );
    }

    const bridgeDir = asString(src.anime.bridgeDir);
    if (bridgeDir !== undefined) {
      resolved.anime.bridgeDir = normalizeExternalProfilePath(bridgeDir);
    } else if (src.anime.bridgeDir !== undefined) {
      warn('anime.bridgeDir', src.anime.bridgeDir, resolved.anime.bridgeDir, 'Expected string.');
    }

    const preferredQuality = asString(src.anime.preferredQuality);
    if (preferredQuality !== undefined) {
      resolved.anime.preferredQuality = preferredQuality.trim();
    } else if (src.anime.preferredQuality !== undefined) {
      warn(
        'anime.preferredQuality',
        src.anime.preferredQuality,
        resolved.anime.preferredQuality,
        'Expected string.',
      );
    }

    if (Array.isArray(src.anime.repos)) {
      const repos: string[] = [];
      for (const entry of src.anime.repos) {
        const url = asString(entry)?.trim();
        // Rejecting here rather than at fetch time keeps a bad entry from
        // silently doing nothing, and blocks plain http downgrades.
        if (url !== undefined && isValidRepoUrl(url)) {
          if (!repos.includes(url)) repos.push(url);
        } else {
          warn('anime.repos[]', entry, null, 'Expected an https URL to a .json index file.');
        }
      }
      resolved.anime.repos = repos;
    } else if (src.anime.repos !== undefined) {
      warn('anime.repos', src.anime.repos, resolved.anime.repos, 'Expected array of strings.');
    }
  } else if (src.anime !== undefined) {
    warn('anime', src.anime, resolved.anime, 'Expected object.');
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

    if (Array.isArray(src.jellyfin.recentServers)) {
      const seenRecentServers = new Set<string>();
      resolved.jellyfin.recentServers = src.jellyfin.recentServers
        .filter((item): item is string => typeof item === 'string')
        .map((item) => item.trim().replace(/\/+$/, ''))
        .filter((item) => {
          if (!item || seenRecentServers.has(item)) return false;
          seenRecentServers.add(item);
          return true;
        })
        .slice(0, 5);
    } else if (src.jellyfin.recentServers !== undefined) {
      warn(
        'jellyfin.recentServers',
        src.jellyfin.recentServers,
        resolved.jellyfin.recentServers,
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
