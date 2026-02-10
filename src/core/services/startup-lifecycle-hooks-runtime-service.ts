import { AppLifecycleDepsRuntimeOptions } from "./app-lifecycle-deps-runtime-service";
import {
  AppReadyRuntimeDeps,
  runAppReadyRuntimeService,
} from "./app-ready-runtime-service";
import {
  AppShutdownRuntimeDeps,
  runAppShutdownRuntimeService,
} from "./app-shutdown-runtime-service";

type StartupLifecycleHookDeps = Pick<
  AppLifecycleDepsRuntimeOptions,
  "onReady" | "onWillQuitCleanup" | "shouldRestoreWindowsOnActivate" | "restoreWindowsOnActivate"
>;

export interface StartupLifecycleHooksRuntimeOptions {
  appReadyDeps: AppReadyRuntimeDeps;
  appShutdownDeps: AppShutdownRuntimeDeps;
  shouldRestoreWindowsOnActivate: () => boolean;
  restoreWindowsOnActivate: () => void;
}

export function createStartupLifecycleHooksRuntimeService(
  options: StartupLifecycleHooksRuntimeOptions,
): StartupLifecycleHookDeps {
  return {
    onReady: async () => {
      await runAppReadyRuntimeService(options.appReadyDeps);
    },
    onWillQuitCleanup: () => {
      runAppShutdownRuntimeService(options.appShutdownDeps);
    },
    shouldRestoreWindowsOnActivate: options.shouldRestoreWindowsOnActivate,
    restoreWindowsOnActivate: options.restoreWindowsOnActivate,
  };
}
