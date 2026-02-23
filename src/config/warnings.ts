import { ConfigValidationWarning } from '../types';

export interface WarningCollector {
  warnings: ConfigValidationWarning[];
  warn(path: string, value: unknown, fallback: unknown, message: string): void;
}

export function createWarningCollector(): WarningCollector {
  const warnings: ConfigValidationWarning[] = [];
  const warn = (path: string, value: unknown, fallback: unknown, message: string): void => {
    warnings.push({
      path,
      value,
      fallback,
      message,
    });
  };
  return { warnings, warn };
}
