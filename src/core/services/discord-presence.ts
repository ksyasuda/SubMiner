import type { ResolvedConfig } from '../../types';

export interface DiscordPresenceSnapshot {
  mediaTitle: string | null;
  mediaPath: string | null;
  subtitleText: string;
  paused: boolean | null;
  connected: boolean;
  sessionStartedAtMs: number;
}

type DiscordPresenceConfig = ResolvedConfig['discordPresence'];

export interface DiscordActivityPayload {
  details: string;
  state: string;
  startTimestamp: number;
  largeImageKey?: string;
  largeImageText?: string;
  smallImageKey?: string;
  smallImageText?: string;
  buttons?: Array<{ label: string; url: string }>;
}

type DiscordClient = {
  login: () => Promise<void>;
  setActivity: (activity: DiscordActivityPayload) => Promise<void>;
  clearActivity: () => Promise<void>;
  destroy: () => void;
};

type TimeoutLike = ReturnType<typeof setTimeout>;

function trimField(value: string, maxLength = 128): string {
  if (value.length <= maxLength) return value;
  return `${value.slice(0, Math.max(0, maxLength - 1))}…`;
}

function sanitizeText(value: string | null | undefined, fallback: string): string {
  const text = value?.trim();
  if (!text) return fallback;
  return text;
}

function basename(filePath: string | null): string {
  if (!filePath) return '';
  const parts = filePath.split(/[\\/]/);
  return parts[parts.length - 1] ?? '';
}

function interpolate(template: string, values: Record<string, string>): string {
  return template.replace(/\{([a-zA-Z0-9_]+)\}/g, (_match, key: string) => values[key] ?? '');
}

function buildStatus(snapshot: DiscordPresenceSnapshot): string {
  if (!snapshot.connected || !snapshot.mediaPath) return 'Idle';
  if (snapshot.paused) return 'Paused';
  return 'Watching';
}

export function buildDiscordPresenceActivity(
  config: DiscordPresenceConfig,
  snapshot: DiscordPresenceSnapshot,
): DiscordActivityPayload {
  const status = buildStatus(snapshot);
  const title = sanitizeText(snapshot.mediaTitle, basename(snapshot.mediaPath) || 'Unknown media');
  const subtitle = sanitizeText(snapshot.subtitleText, '');
  const file = sanitizeText(basename(snapshot.mediaPath), '');
  const values = {
    title,
    file,
    subtitle,
    status,
  };

  const details = trimField(
    interpolate(config.detailsTemplate, values).trim() || `Watching ${title}`,
  );
  const state = trimField(
    interpolate(config.stateTemplate, values).trim() || `${status} with SubMiner`,
  );

  const activity: DiscordActivityPayload = {
    details,
    state,
    startTimestamp: Math.floor(snapshot.sessionStartedAtMs / 1000),
  };

  if (config.largeImageKey.trim().length > 0) {
    activity.largeImageKey = config.largeImageKey.trim();
  }
  if (config.largeImageText.trim().length > 0) {
    activity.largeImageText = trimField(config.largeImageText.trim());
  }
  if (config.smallImageKey.trim().length > 0) {
    activity.smallImageKey = config.smallImageKey.trim();
  }
  if (config.smallImageText.trim().length > 0) {
    activity.smallImageText = trimField(config.smallImageText.trim());
  }
  if (config.buttonLabel.trim().length > 0 && /^https?:\/\//.test(config.buttonUrl.trim())) {
    activity.buttons = [
      { label: trimField(config.buttonLabel.trim(), 32), url: config.buttonUrl.trim() },
    ];
  }

  return activity;
}

export function createDiscordPresenceService(deps: {
  config: DiscordPresenceConfig;
  createClient: (clientId: string) => DiscordClient;
  now?: () => number;
  setTimeoutFn?: (callback: () => void, delayMs: number) => TimeoutLike;
  clearTimeoutFn?: (timer: TimeoutLike) => void;
  logDebug?: (message: string, meta?: unknown) => void;
}) {
  const now = deps.now ?? (() => Date.now());
  const setTimeoutFn = deps.setTimeoutFn ?? ((callback, delayMs) => setTimeout(callback, delayMs));
  const clearTimeoutFn = deps.clearTimeoutFn ?? ((timer) => clearTimeout(timer));
  const logDebug = deps.logDebug ?? (() => {});

  let client: DiscordClient | null = null;
  let pendingSnapshot: DiscordPresenceSnapshot | null = null;
  let debounceTimer: TimeoutLike | null = null;
  let intervalTimer: TimeoutLike | null = null;
  let lastActivityKey = '';
  let lastSentAtMs = 0;

  async function flush(): Promise<void> {
    if (!client || !pendingSnapshot) return;
    const elapsed = now() - lastSentAtMs;
    if (elapsed < deps.config.updateIntervalMs) {
      const delay = Math.max(0, deps.config.updateIntervalMs - elapsed);
      if (intervalTimer) clearTimeoutFn(intervalTimer);
      intervalTimer = setTimeoutFn(() => {
        void flush();
      }, delay);
      return;
    }

    const payload = buildDiscordPresenceActivity(deps.config, pendingSnapshot);
    const activityKey = JSON.stringify(payload);
    if (activityKey === lastActivityKey) return;

    try {
      await client.setActivity(payload);
      lastSentAtMs = now();
      lastActivityKey = activityKey;
    } catch (error) {
      logDebug('[discord-presence] failed to set activity', error);
    }
  }

  function scheduleFlush(snapshot: DiscordPresenceSnapshot): void {
    pendingSnapshot = snapshot;
    if (debounceTimer) {
      clearTimeoutFn(debounceTimer);
    }
    debounceTimer = setTimeoutFn(() => {
      debounceTimer = null;
      void flush();
    }, deps.config.debounceMs);
  }

  return {
    async start(): Promise<void> {
      if (!deps.config.enabled) return;
      const clientId = deps.config.clientId.trim();
      if (!clientId) {
        logDebug('[discord-presence] enabled but clientId missing; skipping start');
        return;
      }
      try {
        client = deps.createClient(clientId);
        await client.login();
      } catch (error) {
        client = null;
        logDebug('[discord-presence] login failed', error);
      }
    },
    publish(snapshot: DiscordPresenceSnapshot): void {
      if (!client) return;
      scheduleFlush(snapshot);
    },
    async stop(): Promise<void> {
      if (debounceTimer) {
        clearTimeoutFn(debounceTimer);
        debounceTimer = null;
      }
      if (intervalTimer) {
        clearTimeoutFn(intervalTimer);
        intervalTimer = null;
      }
      pendingSnapshot = null;
      lastActivityKey = '';
      lastSentAtMs = 0;
      if (!client) return;
      try {
        await client.clearActivity();
      } catch (error) {
        logDebug('[discord-presence] clear activity failed', error);
      }
      client.destroy();
      client = null;
    },
  };
}
