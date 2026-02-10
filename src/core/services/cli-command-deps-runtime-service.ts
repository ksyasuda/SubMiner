import { CliCommandServiceDeps } from "./cli-command-service";

interface MpvClientLike {
  setSocketPath: (socketPath: string) => void;
  connect: () => void;
}

interface TexthookerServiceLike {
  isRunning: () => boolean;
  start: (port: number) => void;
}

interface MpvCliRuntime {
  getSocketPath: () => string;
  setSocketPath: (socketPath: string) => void;
  getClient: () => MpvClientLike | null;
  showOsd: (text: string) => void;
}

interface TexthookerCliRuntime {
  service: TexthookerServiceLike;
  getPort: () => number;
  setPort: (port: number) => void;
  shouldOpenBrowser: () => boolean;
  openInBrowser: (url: string) => void;
}

interface OverlayCliRuntime {
  isInitialized: () => boolean;
  initialize: () => void;
  toggleVisible: () => void;
  toggleInvisible: () => void;
  setVisible: (visible: boolean) => void;
  setInvisible: (visible: boolean) => void;
}

interface MiningCliRuntime {
  copyCurrentSubtitle: () => void;
  startPendingMultiCopy: (timeoutMs: number) => void;
  mineSentenceCard: () => Promise<void>;
  startPendingMineSentenceMultiple: (timeoutMs: number) => void;
  updateLastCardFromClipboard: () => Promise<void>;
  triggerFieldGrouping: () => Promise<void>;
  triggerSubsyncFromConfig: () => Promise<void>;
  markLastCardAsAudioCard: () => Promise<void>;
}

interface UiCliRuntime {
  openYomitanSettings: () => void;
  cycleSecondarySubMode: () => void;
  openRuntimeOptionsPalette: () => void;
  printHelp: () => void;
}

interface AppCliRuntime {
  stop: () => void;
  hasMainWindow: () => boolean;
}

export interface CliCommandDepsRuntimeOptions {
  mpv: MpvCliRuntime;
  texthooker: TexthookerCliRuntime;
  overlay: OverlayCliRuntime;
  mining: MiningCliRuntime;
  ui: UiCliRuntime;
  app: AppCliRuntime;
  getMultiCopyTimeoutMs: () => number;
  schedule: (fn: () => void, delayMs: number) => unknown;
  log: (message: string) => void;
  warn: (message: string) => void;
  error: (message: string, err: unknown) => void;
}

export function createCliCommandDepsRuntimeService(
  options: CliCommandDepsRuntimeOptions,
): CliCommandServiceDeps {
  return {
    getMpvSocketPath: options.mpv.getSocketPath,
    setMpvSocketPath: options.mpv.setSocketPath,
    setMpvClientSocketPath: (socketPath) => {
      const client = options.mpv.getClient();
      if (!client) return;
      client.setSocketPath(socketPath);
    },
    hasMpvClient: () => Boolean(options.mpv.getClient()),
    connectMpvClient: () => {
      const client = options.mpv.getClient();
      if (!client) return;
      client.connect();
    },
    isTexthookerRunning: () => options.texthooker.service.isRunning(),
    setTexthookerPort: options.texthooker.setPort,
    getTexthookerPort: options.texthooker.getPort,
    shouldOpenTexthookerBrowser: options.texthooker.shouldOpenBrowser,
    ensureTexthookerRunning: (port) => {
      if (!options.texthooker.service.isRunning()) {
        options.texthooker.service.start(port);
      }
    },
    openTexthookerInBrowser: options.texthooker.openInBrowser,
    stopApp: options.app.stop,
    isOverlayRuntimeInitialized: options.overlay.isInitialized,
    initializeOverlayRuntime: options.overlay.initialize,
    toggleVisibleOverlay: options.overlay.toggleVisible,
    toggleInvisibleOverlay: options.overlay.toggleInvisible,
    openYomitanSettingsDelayed: (delayMs) => {
      options.schedule(() => {
        options.ui.openYomitanSettings();
      }, delayMs);
    },
    setVisibleOverlayVisible: options.overlay.setVisible,
    setInvisibleOverlayVisible: options.overlay.setInvisible,
    copyCurrentSubtitle: options.mining.copyCurrentSubtitle,
    startPendingMultiCopy: options.mining.startPendingMultiCopy,
    mineSentenceCard: options.mining.mineSentenceCard,
    startPendingMineSentenceMultiple:
      options.mining.startPendingMineSentenceMultiple,
    updateLastCardFromClipboard: options.mining.updateLastCardFromClipboard,
    cycleSecondarySubMode: options.ui.cycleSecondarySubMode,
    triggerFieldGrouping: options.mining.triggerFieldGrouping,
    triggerSubsyncFromConfig: options.mining.triggerSubsyncFromConfig,
    markLastCardAsAudioCard: options.mining.markLastCardAsAudioCard,
    openRuntimeOptionsPalette: options.ui.openRuntimeOptionsPalette,
    printHelp: options.ui.printHelp,
    hasMainWindow: options.app.hasMainWindow,
    getMultiCopyTimeoutMs: options.getMultiCopyTimeoutMs,
    showMpvOsd: options.mpv.showOsd,
    log: options.log,
    warn: options.warn,
    error: options.error,
  };
}
