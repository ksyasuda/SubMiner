export interface AnimeBrowserJimakuAutoOpenDeps {
  isEnabled: () => boolean;
  isAnimeBrowserMedia: (mediaPath: string) => boolean;
  getCurrentMediaPath: () => string | null;
  getPlaybackPaused: () => Promise<boolean | null>;
  setPlaybackPaused: (paused: boolean) => void;
  closeAnimeBrowserModal: () => void;
  openJimakuModal: () => Promise<boolean>;
  logWarn: (message: string, error?: unknown) => void;
}

export interface AnimeBrowserJimakuAutoOpen {
  handleMediaPathChange: (mediaPath: string | null) => Promise<void>;
  handleJimakuSubtitleLoaded: () => void;
  handleJimakuModalClosed: () => void;
}

interface ActiveFlow {
  mediaPath: string;
  ownsPause: boolean;
}

/** Coordinates the pause owned by the Anime Browser to Jimaku handoff. */
export function createAnimeBrowserJimakuAutoOpen(
  deps: AnimeBrowserJimakuAutoOpenDeps,
): AnimeBrowserJimakuAutoOpen {
  let activeFlow: ActiveFlow | null = null;
  let lastMediaPath: string | null = null;

  const isCurrent = (flow: ActiveFlow): boolean =>
    activeFlow === flow && deps.getCurrentMediaPath()?.trim() === flow.mediaPath;

  const releaseFlow = (flow: ActiveFlow | null): void => {
    if (!flow || activeFlow !== flow) return;
    activeFlow = null;
    if (flow.ownsPause) {
      deps.setPlaybackPaused(false);
    }
  };

  const handleMediaPathChange = async (mediaPath: string | null): Promise<void> => {
    const normalizedPath = mediaPath?.trim() || null;
    if (normalizedPath === lastMediaPath) return;
    lastMediaPath = normalizedPath;

    if (!normalizedPath || !deps.isEnabled() || !deps.isAnimeBrowserMedia(normalizedPath)) {
      releaseFlow(activeFlow);
      return;
    }

    const flow: ActiveFlow = {
      mediaPath: normalizedPath,
      ownsPause: activeFlow?.ownsPause ?? false,
    };
    activeFlow = flow;

    if (!flow.ownsPause) {
      try {
        const paused = await deps.getPlaybackPaused();
        if (!isCurrent(flow)) return;
        if (paused === false) {
          deps.setPlaybackPaused(true);
          flow.ownsPause = true;
        }
      } catch (error) {
        deps.logWarn('Could not read playback state before opening Jimaku.', error);
      }
    }

    if (!isCurrent(flow)) return;
    deps.closeAnimeBrowserModal();

    try {
      const opened = await deps.openJimakuModal();
      if (isCurrent(flow) && !opened) {
        releaseFlow(flow);
      }
    } catch (error) {
      deps.logWarn('Could not open Jimaku for Anime Browser playback.', error);
      releaseFlow(flow);
    }
  };

  return {
    handleMediaPathChange,
    handleJimakuSubtitleLoaded: () => releaseFlow(activeFlow),
    handleJimakuModalClosed: () => releaseFlow(activeFlow),
  };
}
