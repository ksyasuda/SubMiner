type MpvClientLike = unknown;

export function createApplyJellyfinMpvDefaultsHandler(deps: {
  sendMpvCommandRuntime: (
    client: MpvClientLike,
    command: [string, string, string],
  ) => void;
  jellyfinLangPref: string;
}) {
  return (client: MpvClientLike): void => {
    deps.sendMpvCommandRuntime(client, ['set_property', 'sub-auto', 'fuzzy']);
    deps.sendMpvCommandRuntime(client, ['set_property', 'aid', 'auto']);
    deps.sendMpvCommandRuntime(client, ['set_property', 'sid', 'auto']);
    deps.sendMpvCommandRuntime(client, ['set_property', 'secondary-sid', 'auto']);
    deps.sendMpvCommandRuntime(client, ['set_property', 'secondary-sub-visibility', 'no']);
    deps.sendMpvCommandRuntime(client, ['set_property', 'alang', deps.jellyfinLangPref]);
    deps.sendMpvCommandRuntime(client, ['set_property', 'slang', deps.jellyfinLangPref]);
  };
}

export function createGetDefaultSocketPathHandler(deps: {
  platform: string;
}) {
  return (): string => {
    if (deps.platform === 'win32') {
      return '\\\\.\\pipe\\subminer-socket';
    }
    return '/tmp/subminer-socket';
  };
}
