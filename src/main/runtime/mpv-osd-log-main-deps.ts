import type { createAppendToMpvLogHandler, createShowMpvOsdHandler } from './mpv-osd-log';

type AppendToMpvLogMainDeps = Parameters<typeof createAppendToMpvLogHandler>[0];
type ShowMpvOsdMainDeps = Parameters<typeof createShowMpvOsdHandler>[0];

export function createBuildAppendToMpvLogMainDepsHandler(deps: AppendToMpvLogMainDeps) {
  return (): AppendToMpvLogMainDeps => ({
    logPath: deps.logPath,
    dirname: (targetPath: string) => deps.dirname(targetPath),
    mkdir: (targetPath: string, options: { recursive: boolean }) => deps.mkdir(targetPath, options),
    appendFile: (targetPath: string, data: string, options: { encoding: 'utf8' }) =>
      deps.appendFile(targetPath, data, options),
    now: () => deps.now(),
  });
}

export function createBuildShowMpvOsdMainDepsHandler(deps: ShowMpvOsdMainDeps) {
  return (): ShowMpvOsdMainDeps => ({
    appendToMpvLog: (message: string) => deps.appendToMpvLog(message),
    showMpvOsdRuntime: (mpvClient, text, fallbackLog) =>
      deps.showMpvOsdRuntime(mpvClient, text, fallbackLog),
    getMpvClient: () => deps.getMpvClient(),
    logInfo: (line: string) => deps.logInfo(line),
  });
}
