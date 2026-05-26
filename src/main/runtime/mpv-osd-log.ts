import type { MpvRuntimeClientLike } from '../../core/services/mpv';

export function createAppendToMpvLogHandler(deps: {
  getLogPath: () => string;
  dirname: (targetPath: string) => string;
  mkdir: (targetPath: string, options: { recursive: boolean }) => Promise<void>;
  appendFile: (targetPath: string, data: string, options: { encoding: 'utf8' }) => Promise<void>;
  now: () => Date;
}) {
  const pendingLines: string[] = [];
  let drainPromise: Promise<void> | null = null;

  const drainPendingLines = async (): Promise<void> => {
    while (pendingLines.length > 0) {
      const chunk = pendingLines.splice(0, pendingLines.length).join('');
      const logPath = deps.getLogPath();
      if (!logPath.trim()) {
        continue;
      }
      try {
        await deps.mkdir(deps.dirname(logPath), { recursive: true });
        await deps.appendFile(logPath, chunk, { encoding: 'utf8' });
      } catch {
        // best-effort logging
      }
    }
  };

  const scheduleDrain = (): Promise<void> => {
    if (drainPromise) return drainPromise;
    drainPromise = (async () => {
      try {
        await drainPendingLines();
      } finally {
        drainPromise = null;
      }
    })();
    return drainPromise;
  };

  const appendToMpvLog = (message: string): void => {
    if (!deps.getLogPath().trim()) {
      return;
    }
    pendingLines.push(`[${deps.now().toISOString()}] ${message}\n`);
    void scheduleDrain();
  };

  const flushMpvLog = async (): Promise<void> => {
    while (pendingLines.length > 0 || drainPromise) {
      await scheduleDrain();
    }
  };

  return {
    appendToMpvLog,
    flushMpvLog,
  };
}

export function createShowMpvOsdHandler(deps: {
  appendToMpvLog: (message: string) => void;
  showMpvOsdRuntime: (
    mpvClient: MpvRuntimeClientLike | null,
    text: string,
    fallbackLog: (line: string) => void,
  ) => boolean;
  getMpvClient: () => MpvRuntimeClientLike | null;
  logInfo: (line: string) => void;
}) {
  return (text: string): boolean => {
    deps.appendToMpvLog(`[OSD] ${text}`);
    return deps.showMpvOsdRuntime(deps.getMpvClient(), text, (line) => {
      deps.logInfo(line);
    });
  };
}
