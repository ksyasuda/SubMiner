import {
  createOnWillQuitCleanupHandler,
  createRestoreWindowsOnActivateHandler,
  createShouldRestoreWindowsOnActivateHandler,
} from '../app-lifecycle-actions';
import { createBuildOnWillQuitCleanupDepsHandler } from '../app-lifecycle-main-cleanup';
import {
  createBuildRestoreWindowsOnActivateMainDepsHandler,
  createBuildShouldRestoreWindowsOnActivateMainDepsHandler,
} from '../app-lifecycle-main-activate';
import { registerProtocolUrlHandlers } from '../protocol-url-handlers';
import type { ComposerInputs, ComposerOutputs } from './contracts';

type RegisterProtocolUrlHandlersMainDeps = Parameters<typeof registerProtocolUrlHandlers>[0];
type OnWillQuitCleanupDeps = Parameters<typeof createBuildOnWillQuitCleanupDepsHandler>[0];
type ShouldRestoreWindowsOnActivateMainDeps = Parameters<
  typeof createBuildShouldRestoreWindowsOnActivateMainDepsHandler
>[0];
type RestoreWindowsOnActivateMainDeps = Parameters<
  typeof createBuildRestoreWindowsOnActivateMainDepsHandler
>[0];

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
    createBuildOnWillQuitCleanupDepsHandler(options.onWillQuitCleanupMainDeps)(),
  );
  const shouldRestoreWindowsOnActivateHandler = createShouldRestoreWindowsOnActivateHandler(
    createBuildShouldRestoreWindowsOnActivateMainDepsHandler(
      options.shouldRestoreWindowsOnActivateMainDeps,
    )(),
  );
  const restoreWindowsOnActivateHandler = createRestoreWindowsOnActivateHandler(
    createBuildRestoreWindowsOnActivateMainDepsHandler(options.restoreWindowsOnActivateMainDeps)(),
  );

  return {
    registerProtocolUrlHandlers: () =>
      registerProtocolUrlHandlers(options.registerProtocolUrlHandlersMainDeps),
    onWillQuitCleanup: onWillQuitCleanupHandler,
    shouldRestoreWindowsOnActivate: shouldRestoreWindowsOnActivateHandler,
    restoreWindowsOnActivate: restoreWindowsOnActivateHandler,
  };
}
