import type { MpvRuntimeClientLike } from '../../core/services/mpv';
import { getDefaultMpvSocketPath } from '../../shared/mpv-socket-path';

export function createApplyJellyfinMpvDefaultsHandler(deps: {
  sendMpvCommandRuntime: (client: MpvRuntimeClientLike, command: [string, string, string]) => void;
  jellyfinLangPref: string;
}) {
  return (client: MpvRuntimeClientLike): void => {
    deps.sendMpvCommandRuntime(client, ['set_property', 'sub-auto', 'fuzzy']);
    deps.sendMpvCommandRuntime(client, ['set_property', 'aid', 'auto']);
    deps.sendMpvCommandRuntime(client, ['set_property', 'sid', 'auto']);
    deps.sendMpvCommandRuntime(client, ['set_property', 'secondary-sid', 'auto']);
    deps.sendMpvCommandRuntime(client, ['set_property', 'secondary-sub-visibility', 'no']);
    deps.sendMpvCommandRuntime(client, ['set_property', 'alang', deps.jellyfinLangPref]);
    deps.sendMpvCommandRuntime(client, ['set_property', 'slang', deps.jellyfinLangPref]);
  };
}

export function createGetDefaultSocketPathHandler(deps: { platform: string }) {
  return (): string => {
    return getDefaultMpvSocketPath(deps.platform as NodeJS.Platform);
  };
}
