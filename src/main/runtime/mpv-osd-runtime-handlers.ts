import { createAppendToMpvLogHandler, createShowMpvOsdHandler } from './mpv-osd-log';
import {
  createBuildAppendToMpvLogMainDepsHandler,
  createBuildShowMpvOsdMainDepsHandler,
} from './mpv-osd-log-main-deps';

type AppendToMpvLogMainDeps = Parameters<typeof createBuildAppendToMpvLogMainDepsHandler>[0];
type ShowMpvOsdMainDeps = Parameters<typeof createBuildShowMpvOsdMainDepsHandler>[0];

export function createMpvOsdRuntimeHandlers(deps: {
  appendToMpvLogMainDeps: AppendToMpvLogMainDeps;
  buildShowMpvOsdMainDeps: (appendToMpvLog: (message: string) => void) => ShowMpvOsdMainDeps;
}) {
  const appendToMpvLogMainDeps =
    createBuildAppendToMpvLogMainDepsHandler(deps.appendToMpvLogMainDeps)();
  const appendToMpvLogHandler = createAppendToMpvLogHandler(appendToMpvLogMainDeps);
  const appendToMpvLog = (message: string) => appendToMpvLogHandler(message);

  const showMpvOsdMainDeps = createBuildShowMpvOsdMainDepsHandler(
    deps.buildShowMpvOsdMainDeps(appendToMpvLog),
  )();
  const showMpvOsdHandler = createShowMpvOsdHandler(showMpvOsdMainDeps);
  const showMpvOsd = (text: string) => showMpvOsdHandler(text);

  return {
    appendToMpvLog,
    showMpvOsd,
  };
}
