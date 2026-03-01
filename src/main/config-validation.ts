import type { ConfigValidationWarning } from '../types';

export type StartupFailureHandlers = {
  logError: (details: string) => void;
  showErrorBox: (title: string, details: string) => void;
  quit: () => void;
};

export function formatConfigValue(value: unknown): string {
  if (value === undefined) {
    return 'undefined';
  }

  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

export function buildConfigWarningSummary(
  configPath: string,
  warnings: ConfigValidationWarning[],
): string {
  const lines = [
    `[config] Validation found ${warnings.length} issue(s). File: ${configPath}`,
    ...warnings.map(
      (warning, index) =>
        `[config] ${index + 1}. ${warning.path}: ${warning.message} actual=${formatConfigValue(warning.value)} fallback=${formatConfigValue(warning.fallback)}`,
    ),
  ];
  return lines.join('\n');
}

export function buildConfigWarningNotificationBody(
  configPath: string,
  warnings: ConfigValidationWarning[],
): string {
  const maxLines = 3;
  const maxPathLength = 48;

  const trimPath = (value: string): string =>
    value.length > maxPathLength ? `...${value.slice(-(maxPathLength - 3))}` : value;
  const clippedPath = trimPath(configPath);

  const lines = warnings.slice(0, maxLines).map((warning, index) => {
    const message = `${warning.path}: ${warning.message}`;
    return `${index + 1}. ${message}`;
  });

  const overflow = warnings.length - lines.length;
  if (overflow > 0) {
    lines.push(`+${overflow} more issue(s)`);
  }

  return [
    `${warnings.length} config validation issue(s) detected.`,
    'Defaults were applied where possible.',
    `File: ${clippedPath}`,
    ...lines,
  ].join('\n');
}

export function buildConfigWarningDialogDetails(
  configPath: string,
  warnings: ConfigValidationWarning[],
): string {
  const lines = warnings.map(
    (warning, index) =>
      `${index + 1}. ${warning.path}: ${warning.message} actual=${formatConfigValue(warning.value)} fallback=${formatConfigValue(warning.fallback)}`,
  );

  return [
    'SubMiner detected config validation issues.',
    `File: ${configPath}`,
    '',
    ...lines,
    '',
    'Defaults were applied where possible.',
  ].join('\n');
}

export function buildConfigParseErrorDetails(configPath: string, parseError: string): string {
  return [
    'Failed to parse config file at:',
    configPath,
    '',
    `Error: ${parseError}`,
    '',
    'Fix the config file and restart SubMiner.',
  ].join('\n');
}

export function failStartupFromConfig(
  title: string,
  details: string,
  handlers: StartupFailureHandlers,
): never {
  handlers.logError(details);
  handlers.showErrorBox(title, details);
  process.exitCode = 1;
  handlers.quit();
  throw new Error(details);
}
