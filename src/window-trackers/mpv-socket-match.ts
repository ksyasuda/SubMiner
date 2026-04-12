import type { MpvPollResult } from './win32';

function escapeRegex(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function matchesMpvSocketPathInCommandLine(
  commandLine: string,
  targetSocketPath: string,
): boolean {
  if (!commandLine || !targetSocketPath) {
    return false;
  }

  const escapedSocketPath = escapeRegex(targetSocketPath);
  return new RegExp(
    `(?:^|\\s)--input-ipc-server(?:=|\\s+)(?:"${escapedSocketPath}"|${escapedSocketPath})(?=\\s|$)`,
    'i',
  ).test(commandLine);
}

export function filterMpvPollResultBySocketPath(
  result: MpvPollResult,
  targetSocketPath?: string | null,
): MpvPollResult {
  if (!targetSocketPath) {
    return result;
  }

  const matches = result.matches.filter(
    (match) =>
      typeof match.commandLine === 'string' &&
      matchesMpvSocketPathInCommandLine(match.commandLine, targetSocketPath),
  );

  return {
    matches,
    focusState: matches.some((match) => match.isForeground),
    windowState: matches.length > 0 ? 'visible' : 'not-found',
  };
}
