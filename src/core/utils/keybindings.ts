import { Config, Keybinding } from '../../types';

export function resolveKeybindings(config: Config, defaultKeybindings: Keybinding[]): Keybinding[] {
  const userBindings = config.keybindings || [];
  const bindingMap = new Map<string, (string | number)[] | null>();

  for (const binding of defaultKeybindings) {
    bindingMap.set(binding.key, binding.command);
  }

  for (const binding of userBindings) {
    if (binding.command === null) {
      bindingMap.delete(binding.key);
    } else {
      bindingMap.set(binding.key, binding.command);
    }
  }

  const keybindings: Keybinding[] = [];
  for (const [key, command] of bindingMap) {
    if (command !== null) {
      keybindings.push({ key, command });
    }
  }

  return keybindings;
}
