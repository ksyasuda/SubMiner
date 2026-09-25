import type { Keybinding } from '../../types';
import { parseSessionBindingKey } from '../../core/services/session-bindings';
import { parseMpvInputBindingKeys } from '../../shared/mpv-input-bindings';
import type { MpvInputBindingsSnapshot } from '../../types/session-bindings';

export async function readMpvInputBindings(deps: {
  getMpvClient: () => {
    connected: boolean;
    requestProperty: (name: string) => Promise<unknown>;
  } | null;
  getConfiguredKeybindings: () => Keybinding[];
  platform: 'darwin' | 'win32' | 'linux';
}): Promise<MpvInputBindingsSnapshot> {
  const blockedKeys = deps.getConfiguredKeybindings().flatMap((binding) => {
    const { key } = parseSessionBindingKey(binding.key, deps.platform);
    return key ? [key] : [];
  });
  const client = deps.getMpvClient();
  if (!client?.connected) return { keys: [], blockedKeys };
  try {
    const value = await client.requestProperty('input-bindings');
    return {
      keys:
        client === deps.getMpvClient() && client.connected ? parseMpvInputBindingKeys(value) : [],
      blockedKeys,
    };
  } catch {
    // Older mpv versions and disconnected sessions retain SubMiner's controls.
    return { keys: [], blockedKeys };
  }
}
