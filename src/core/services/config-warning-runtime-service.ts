import { ConfigValidationWarning } from "../../types";

export function formatConfigWarningRuntimeService(
  warning: ConfigValidationWarning,
): string {
  return `[config] ${warning.path}: ${warning.message} value=${JSON.stringify(warning.value)} fallback=${JSON.stringify(warning.fallback)}`;
}

export function logConfigWarningRuntimeService(
  warning: ConfigValidationWarning,
  log: (message: string) => void,
): void {
  log(formatConfigWarningRuntimeService(warning));
}
