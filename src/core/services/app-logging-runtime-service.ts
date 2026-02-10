import { ConfigValidationWarning } from "../../types";
import { logConfigWarningRuntimeService } from "./config-warning-runtime-service";

export interface AppLoggingRuntime {
  logInfo: (message: string) => void;
  logWarning: (message: string) => void;
  logNoRunningInstance: () => void;
  logConfigWarning: (warning: ConfigValidationWarning) => void;
}

export function createAppLoggingRuntimeService(
  logger: Pick<Console, "log" | "warn" | "error"> = console,
): AppLoggingRuntime {
  return {
    logInfo: (message) => {
      logger.log(message);
    },
    logWarning: (message) => {
      logger.warn(message);
    },
    logNoRunningInstance: () => {
      logger.error("No running instance. Use --start to launch the app.");
    },
    logConfigWarning: (warning) => {
      logConfigWarningRuntimeService(warning, (line) => logger.warn(line));
    },
  };
}
