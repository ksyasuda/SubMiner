import { appendLogLine, resolveDefaultLogFilePath } from '../shared/log-files';
import { i18n } from '../i18n/index.js';

export type FatalErrorReportOptions = {
  title: string;
  context: string;
};

export type FatalErrorReporterDeps = {
  showErrorBox: (title: string, details: string) => void;
  consoleError?: (message: string, error?: unknown) => void;
  appendLogLine?: (line: string) => void;
  resolveLogFilePath?: () => string;
  now?: () => Date;
};

export type FatalErrorReporter = (
  error: unknown,
  options?: Partial<FatalErrorReportOptions>,
) => void;

const DEFAULT_TITLE = i18n.t('dialog.crashed');
const DEFAULT_CONTEXT = i18n.t('dialog.fatalContext');

let fatalErrorReported = false;

function pad(value: number): string {
  return String(value).padStart(2, '0');
}

function formatTimestamp(date: Date): string {
  return [
    date.getFullYear(),
    '-',
    pad(date.getMonth() + 1),
    '-',
    pad(date.getDate()),
    ' ',
    pad(date.getHours()),
    ':',
    pad(date.getMinutes()),
    ':',
    pad(date.getSeconds()),
  ].join('');
}

function stringifyUnknownError(error: unknown): string {
  if (error instanceof Error) {
    return error.stack || error.message || error.name;
  }
  if (typeof error === 'string') {
    return error;
  }
  try {
    return JSON.stringify(error) ?? String(error);
  } catch {
    return String(error);
  }
}

export function buildFatalErrorDetails(options: {
  context: string;
  error: unknown;
  logFilePath: string;
}): string {
  return [
    options.context,
    '',
    stringifyUnknownError(options.error),
    '',
    `Log file: ${options.logFilePath}`,
  ].join('\n');
}

export function createFatalErrorReporter(deps: FatalErrorReporterDeps): FatalErrorReporter {
  return (error, options = {}) => {
    if (fatalErrorReported) {
      return;
    }
    fatalErrorReported = true;

    const title = options.title ?? DEFAULT_TITLE;
    const context = options.context ?? DEFAULT_CONTEXT;
    const logFilePath = deps.resolveLogFilePath?.() ?? resolveDefaultLogFilePath('app');
    const details = buildFatalErrorDetails({ context, error, logFilePath });
    const timestamp = formatTimestamp(deps.now?.() ?? new Date());
    const line = `[subminer] - ${timestamp} - ERROR - [main:fatal] ${details.replace(/\r?\n/g, ' | ')}`;

    try {
      (deps.appendLogLine ?? ((entry: string) => appendLogLine(logFilePath, entry)))(line);
    } catch {
      // Fatal reporting must never throw while handling the original failure.
    }

    try {
      deps.consoleError?.(line, error);
    } catch {
      // ignore console sink failures
    }

    try {
      deps.showErrorBox(title, details);
    } catch {
      // If native dialogs are unavailable, the file log above is still the source of truth.
    }
  };
}

export function registerFatalErrorHandlers(deps: {
  reportFatalError: FatalErrorReporter;
  exit: (code: number) => void;
}): void {
  process.on('uncaughtException', (error) => {
    deps.reportFatalError(error, {
      title: i18n.t('dialog.crashed'),
      context: i18n.t('dialog.fatalContext'),
    });
    deps.exit(1);
  });

  process.on('unhandledRejection', (reason) => {
    deps.reportFatalError(reason, {
      title: i18n.t('dialog.crashed'),
      context: i18n.t('dialog.rejectionContext'),
    });
    deps.exit(1);
  });
}

export function resetFatalErrorReporterForTests(): void {
  fatalErrorReported = false;
}
