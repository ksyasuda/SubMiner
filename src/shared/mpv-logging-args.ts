export type SharedLogLevel = 'debug' | 'info' | 'warn' | 'error';

function hasOption(args: readonly string[], option: string): boolean {
  return args.some((arg) => arg === option || arg.startsWith(`${option}=`));
}

export function buildMpvMsgLevel(logLevel: SharedLogLevel): string {
  return `all=warn,subminer=${logLevel}`;
}

export function buildMpvLoggingArgs(
  logLevel: SharedLogLevel,
  logPath: string,
  existingArgs: readonly string[] = [],
): string[] {
  if (!logPath.trim()) {
    return [];
  }
  const args: string[] = [];
  if (!hasOption(existingArgs, '--log-file')) {
    args.push(`--log-file=${logPath}`);
  }
  if (!hasOption(existingArgs, '--msg-level')) {
    args.push(`--msg-level=${buildMpvMsgLevel(logLevel)}`);
  }
  return args;
}
