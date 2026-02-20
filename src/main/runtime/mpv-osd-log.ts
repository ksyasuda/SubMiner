import type { MpvRuntimeClientLike } from '../../core/services/mpv';

export function createAppendToMpvLogHandler(deps: {
  logPath: string;
  dirname: (targetPath: string) => string;
  mkdirSync: (targetPath: string, options: { recursive: boolean }) => void;
  appendFileSync: (
    targetPath: string,
    data: string,
    options: { encoding: 'utf8' },
  ) => void;
  now: () => Date;
}) {
  return (message: string): void => {
    try {
      deps.mkdirSync(deps.dirname(deps.logPath), { recursive: true });
      deps.appendFileSync(deps.logPath, `[${deps.now().toISOString()}] ${message}\n`, {
        encoding: 'utf8',
      });
    } catch {
      // best-effort logging
    }
  };
}

export function createShowMpvOsdHandler(deps: {
  appendToMpvLog: (message: string) => void;
  showMpvOsdRuntime: (
    mpvClient: MpvRuntimeClientLike | null,
    text: string,
    fallbackLog: (line: string) => void,
  ) => void;
  getMpvClient: () => MpvRuntimeClientLike | null;
  logInfo: (line: string) => void;
}) {
  return (text: string): void => {
    deps.appendToMpvLog(`[OSD] ${text}`);
    deps.showMpvOsdRuntime(deps.getMpvClient(), text, (line) => {
      deps.logInfo(line);
    });
  };
}
