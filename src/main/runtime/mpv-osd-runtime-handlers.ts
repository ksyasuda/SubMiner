import { createAppendToMpvLogHandler, createShowMpvOsdHandler } from './mpv-osd-log';

type AppendToMpvLogMainDeps = Parameters<typeof createAppendToMpvLogHandler>[0];
type ShowMpvOsdMainDeps = Parameters<typeof createShowMpvOsdHandler>[0];

export function createMpvOsdRuntimeHandlers(deps: {
  appendToMpvLogMainDeps: AppendToMpvLogMainDeps;
  buildShowMpvOsdMainDeps: (appendToMpvLog: (message: string) => void) => ShowMpvOsdMainDeps;
}) {
  const appendToMpvLogRuntime = createAppendToMpvLogHandler(deps.appendToMpvLogMainDeps);
  const appendToMpvLog = (message: string) => appendToMpvLogRuntime.appendToMpvLog(message);
  const flushMpvLog = async () => appendToMpvLogRuntime.flushMpvLog();

  const showMpvOsdHandler = createShowMpvOsdHandler(deps.buildShowMpvOsdMainDeps(appendToMpvLog));
  const showMpvOsd = (text: string) => showMpvOsdHandler(text);

  return {
    appendToMpvLog,
    flushMpvLog,
    showMpvOsd,
  };
}
