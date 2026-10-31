import {
  createOnWillQuitCleanupHandler,
  createRestoreWindowsOnActivateHandler,
  createShouldRestoreWindowsOnActivateHandler,
} from '../app-lifecycle-actions';
import { createOnWillQuitCleanupDeps } from '../app-lifecycle-main-cleanup';
import { registerProtocolUrlHandlers } from '../protocol-url-handlers';
import type { ComposerInputs, ComposerOutputs } from './contracts';

type RegisterProtocolUrlHandlersMainDeps = Parameters<typeof registerProtocolUrlHandlers>[0];
type OnWillQuitCleanupDeps = Parameters<typeof createOnWillQuitCleanupDeps>[0];
type ShouldRestoreWindowsOnActivateMainDeps = Parameters<
  typeof createShouldRestoreWindowsOnActivateHandler
>[0];
type RestoreWindowsOnActivateMainDeps = Parameters<typeof createRestoreWindowsOnActivateHandler>[0];

export type StartupLifecycleComposerOptions = ComposerInputs<{
  registerProtocolUrlHandlersMainDeps: RegisterProtocolUrlHandlersMainDeps;
  onWillQuitCleanupMainDeps: OnWillQuitCleanupDeps;
  shouldRestoreWindowsOnActivateMainDeps: ShouldRestoreWindowsOnActivateMainDeps;
  restoreWindowsOnActivateMainDeps: RestoreWindowsOnActivateMainDeps;
}>;

export type StartupLifecycleComposerResult = ComposerOutputs<{
  registerProtocolUrlHandlers: () => void;
  onWillQuitCleanup: () => Promise<void>;
  shouldRestoreWindowsOnActivate: () => boolean;
  restoreWindowsOnActivate: () => void;
}>;

export function composeStartupLifecycleHandlers(
  options: StartupLifecycleComposerOptions,
): StartupLifecycleComposerResult {
  const onWillQuitCleanupHandler = createOnWillQuitCleanupHandler(
    createOnWillQuitCleanupDeps(options.onWillQuitCleanupMainDeps),
  );
  const shouldRestoreWindowsOnActivateHandler = createShouldRestoreWindowsOnActivateHandler(
    options.shouldRestoreWindowsOnActivateMainDeps,
  );
  const restoreWindowsOnActivateHandler = createRestoreWindowsOnActivateHandler(
    options.restoreWindowsOnActivateMainDeps,
  );

  return {
    registerProtocolUrlHandlers: () =>
      registerProtocolUrlHandlers(options.registerProtocolUrlHandlersMainDeps),
    onWillQuitCleanup: onWillQuitCleanupHandler,
    shouldRestoreWindowsOnActivate: shouldRestoreWindowsOnActivateHandler,
    restoreWindowsOnActivate: restoreWindowsOnActivateHandler,
  };
}
