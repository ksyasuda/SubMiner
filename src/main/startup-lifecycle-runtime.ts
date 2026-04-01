import {
  composeStartupLifecycleHandlers,
  type StartupLifecycleComposerOptions,
} from './runtime/composers';

export interface StartupLifecycleRuntimeInput {
  protocolUrl: StartupLifecycleComposerOptions['registerProtocolUrlHandlersMainDeps'];
  cleanup: StartupLifecycleComposerOptions['onWillQuitCleanupMainDeps'];
  shouldRestoreWindowsOnActivate: StartupLifecycleComposerOptions['shouldRestoreWindowsOnActivateMainDeps'];
  restoreWindowsOnActivate: StartupLifecycleComposerOptions['restoreWindowsOnActivateMainDeps'];
}

export interface StartupLifecycleRuntime {
  registerProtocolUrlHandlers: () => void;
  onWillQuitCleanup: () => void;
  shouldRestoreWindowsOnActivate: () => boolean;
  restoreWindowsOnActivate: () => void;
}

export function createStartupLifecycleRuntime(
  input: StartupLifecycleRuntimeInput,
): StartupLifecycleRuntime {
  const {
    registerProtocolUrlHandlers,
    onWillQuitCleanup,
    shouldRestoreWindowsOnActivate,
    restoreWindowsOnActivate,
  } = composeStartupLifecycleHandlers({
    registerProtocolUrlHandlersMainDeps: input.protocolUrl,
    onWillQuitCleanupMainDeps: input.cleanup,
    shouldRestoreWindowsOnActivateMainDeps: input.shouldRestoreWindowsOnActivate,
    restoreWindowsOnActivateMainDeps: input.restoreWindowsOnActivate,
  });

  return {
    registerProtocolUrlHandlers,
    onWillQuitCleanup,
    shouldRestoreWindowsOnActivate,
    restoreWindowsOnActivate,
  };
}
