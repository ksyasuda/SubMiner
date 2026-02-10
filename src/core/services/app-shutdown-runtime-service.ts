export interface AppShutdownRuntimeDeps {
  unregisterAllGlobalShortcuts: () => void;
  stopSubtitleWebsocket: () => void;
  stopTexthookerService: () => void;
  destroyYomitanParserWindow: () => void;
  clearYomitanParserPromises: () => void;
  stopWindowTracker: () => void;
  destroyMpvSocket: () => void;
  clearReconnectTimer: () => void;
  destroySubtitleTimingTracker: () => void;
  destroyAnkiIntegration: () => void;
}

export function runAppShutdownRuntimeService(
  deps: AppShutdownRuntimeDeps,
): void {
  deps.unregisterAllGlobalShortcuts();
  deps.stopSubtitleWebsocket();
  deps.stopTexthookerService();
  deps.destroyYomitanParserWindow();
  deps.clearYomitanParserPromises();
  deps.stopWindowTracker();
  deps.destroyMpvSocket();
  deps.clearReconnectTimer();
  deps.destroySubtitleTimingTracker();
  deps.destroyAnkiIntegration();
}
