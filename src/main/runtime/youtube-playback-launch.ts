import { isYoutubeMediaPath } from './youtube-playback';

type YoutubePlaybackLaunchInput = {
  url: string;
  timeoutMs?: number;
  pollIntervalMs?: number;
};

type YoutubePlaybackLaunchDeps = {
  requestPath: () => Promise<string | null>;
  requestProperty?: (name: string) => Promise<unknown>;
  sendMpvCommand: (command: Array<string>) => void;
  wait: (ms: number) => Promise<void>;
  now?: () => number;
};

function normalizePath(value: string | null | undefined): string {
  if (typeof value !== 'string') return '';
  return value.trim();
}

function extractYoutubeVideoId(url: string): string | null {
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    return null;
  }

  const host = parsed.hostname.toLowerCase();
  const path = parsed.pathname.replace(/^\/+/, '');

  if (host === 'youtu.be' || host.endsWith('.youtu.be')) {
    const id = path.split('/')[0]?.trim() || '';
    return id || null;
  }

  const youtubeHost =
    host === 'youtube.com' ||
    host.endsWith('.youtube.com') ||
    host === 'youtube-nocookie.com' ||
    host.endsWith('.youtube-nocookie.com');
  if (!youtubeHost) {
    return null;
  }

  if (parsed.pathname === '/watch') {
    const id = parsed.searchParams.get('v')?.trim() || '';
    return id || null;
  }

  if (path.startsWith('shorts/') || path.startsWith('embed/')) {
    const id = path.split('/')[1]?.trim() || '';
    return id || null;
  }

  return null;
}

function targetsSameYoutubeVideo(currentPath: string, targetUrl: string): boolean {
  const currentId = extractYoutubeVideoId(currentPath);
  const targetId = extractYoutubeVideoId(targetUrl);
  if (!currentId || !targetId) return false;
  return currentId === targetId;
}

function pathMatchesYoutubeTarget(currentPath: string, targetUrl: string): boolean {
  if (!currentPath) return false;
  if (currentPath === targetUrl) return true;
  return targetsSameYoutubeVideo(currentPath, targetUrl);
}

function hasPlayableMediaTracks(trackListRaw: unknown): boolean {
  if (!Array.isArray(trackListRaw)) return false;
  return trackListRaw.some((track) => {
    if (!track || typeof track !== 'object') return false;
    const type = String((track as Record<string, unknown>).type || '')
      .trim()
      .toLowerCase();
    return type === 'video' || type === 'audio';
  });
}

function sendPlaybackPrepCommands(sendMpvCommand: (command: Array<string>) => void): void {
  sendMpvCommand(['set_property', 'pause', 'yes']);
  sendMpvCommand(['set_property', 'sub-auto', 'no']);
  sendMpvCommand(['set_property', 'sid', 'no']);
  sendMpvCommand(['set_property', 'secondary-sid', 'no']);
}

export function createPrepareYoutubePlaybackInMpvHandler(deps: YoutubePlaybackLaunchDeps) {
  const now = deps.now ?? (() => Date.now());
  return async (input: YoutubePlaybackLaunchInput): Promise<boolean> => {
    const targetUrl = input.url.trim();
    if (!targetUrl) return false;

    const timeoutMs = Math.max(200, input.timeoutMs ?? 5000);
    const pollIntervalMs = Math.max(25, input.pollIntervalMs ?? 100);

    let previousPath = '';
    try {
      previousPath = normalizePath(await deps.requestPath());
    } catch {
      // Ignore transient path request failures and continue with bootstrap commands.
    }

    sendPlaybackPrepCommands(deps.sendMpvCommand);

    const alreadyTarget = pathMatchesYoutubeTarget(previousPath, targetUrl);
    if (alreadyTarget) {
      if (!deps.requestProperty) {
        return true;
      }
      try {
        const trackList = await deps.requestProperty('track-list');
        if (hasPlayableMediaTracks(trackList)) {
          return true;
        }
      } catch {
        // Keep polling; mpv can report the target path before tracks are ready.
      }
    } else {
      deps.sendMpvCommand(['loadfile', targetUrl, 'replace']);
    }

    const deadline = now() + timeoutMs;
    while (now() < deadline) {
      await deps.wait(pollIntervalMs);
      let currentPath = '';
      try {
        currentPath = normalizePath(await deps.requestPath());
      } catch {
        continue;
      }
      if (!currentPath) continue;
      if (pathMatchesYoutubeTarget(currentPath, targetUrl)) {
        if (!deps.requestProperty) {
          return true;
        }
        try {
          const trackList = await deps.requestProperty('track-list');
          if (hasPlayableMediaTracks(trackList)) {
            return true;
          }
        } catch {
          // Continue polling until media tracks are actually available.
        }
      }
      const pathDiffersFromInitial = currentPath !== previousPath;
      const matchesChangedTarget =
        currentPath === targetUrl ||
        (isYoutubeMediaPath(currentPath) &&
          isYoutubeMediaPath(targetUrl) &&
          pathMatchesYoutubeTarget(currentPath, targetUrl));
      if (pathDiffersFromInitial && matchesChangedTarget) {
        if (deps.requestProperty) {
          try {
            const trackList = await deps.requestProperty('track-list');
            if (hasPlayableMediaTracks(trackList)) {
              return true;
            }
          } catch {
            // Continue polling until media tracks are actually available.
          }
        }
      }
    }

    return false;
  };
}
