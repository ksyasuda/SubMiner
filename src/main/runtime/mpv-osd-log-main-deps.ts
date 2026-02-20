export function createBuildAppendToMpvLogMainDepsHandler(deps: {
  logPath: string;
  dirname: (targetPath: string) => string;
  mkdirSync: (targetPath: string, options: { recursive: boolean }) => void;
  appendFileSync: (targetPath: string, data: string, options: { encoding: 'utf8' }) => void;
  now: () => Date;
}) {
  return () => ({
    logPath: deps.logPath,
    dirname: (targetPath: string) => deps.dirname(targetPath),
    mkdirSync: (targetPath: string, options: { recursive: boolean }) =>
      deps.mkdirSync(targetPath, options),
    appendFileSync: (targetPath: string, data: string, options: { encoding: 'utf8' }) =>
      deps.appendFileSync(targetPath, data, options),
    now: () => deps.now(),
  });
}

export function createBuildShowMpvOsdMainDepsHandler(deps: {
  appendToMpvLog: (message: string) => void;
  showMpvOsdRuntime: (
    mpvClient: unknown | null,
    text: string,
    fallbackLog: (line: string) => void,
  ) => void;
  getMpvClient: () => unknown | null;
  logInfo: (line: string) => void;
}) {
  return () => ({
    appendToMpvLog: (message: string) => deps.appendToMpvLog(message),
    showMpvOsdRuntime: (
      mpvClient: unknown | null,
      text: string,
      fallbackLog: (line: string) => void,
    ) => deps.showMpvOsdRuntime(mpvClient, text, fallbackLog),
    getMpvClient: () => deps.getMpvClient() as never,
    logInfo: (line: string) => deps.logInfo(line),
  });
}
