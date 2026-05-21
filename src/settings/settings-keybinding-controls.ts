import { DEFAULT_KEYBINDINGS } from '../config/definitions/shared';
import type { ConfigSettingsField } from '../types/settings';
import {
  buildMpvKeybindingConfigValue,
  createMpvKeybindingRows,
  keyboardEventToConfigKey,
  parseMpvCommandText,
  type KeyInputMode,
  type MpvKeybindingRow,
} from './key-input';
import type { SettingsControlContext } from './settings-control-context';
import { createElement } from './settings-control-dom';

let activeKeyLearningStop: (() => void) | null = null;
let requestRender = (): void => undefined;

export function configureKeybindingControls(options: { requestRender: () => void }): void {
  requestRender = options.requestRender;
}

function startKeyLearning(
  button: HTMLButtonElement,
  mode: KeyInputMode,
  onValue: (value: string) => void,
): void {
  activeKeyLearningStop?.();
  const previousText = button.textContent ?? '';
  button.textContent = 'Press Keys...';
  button.classList.add('learning');
  let onKeyDown: (event: KeyboardEvent) => void;
  let onBlur: () => void;
  let onMouseDown: (event: MouseEvent) => void;

  const stop = (): void => {
    window.removeEventListener('keydown', onKeyDown, true);
    window.removeEventListener('blur', onBlur, true);
    window.removeEventListener('mousedown', onMouseDown, true);
    button.classList.remove('learning');
    if (button.textContent === 'Press Keys...') {
      button.textContent = previousText;
    }
    if (activeKeyLearningStop === stop) {
      activeKeyLearningStop = null;
    }
  };

  onKeyDown = (event: KeyboardEvent): void => {
    if (event.key === 'Escape') {
      stop();
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    const next = keyboardEventToConfigKey(event, mode);
    if (!next) return;
    stop();
    onValue(next);
  };
  onBlur = (): void => stop();
  onMouseDown = (event: MouseEvent): void => {
    if (event.target !== button) {
      stop();
    }
  };

  window.addEventListener('keydown', onKeyDown, true);
  window.addEventListener('blur', onBlur, true);
  window.addEventListener('mousedown', onMouseDown, true);
  activeKeyLearningStop = stop;
}

function renderKeyLearnButton(
  value: string,
  mode: KeyInputMode,
  onValue: (value: string) => void,
): HTMLButtonElement {
  const button = createElement('button', 'key-learn-button') as HTMLButtonElement;
  button.type = 'button';
  button.textContent = value || 'Unset';
  button.addEventListener('click', () =>
    startKeyLearning(button, mode, (next) => {
      button.textContent = next;
      onValue(next);
    }),
  );
  return button;
}

export function renderKeyboardInput(
  context: SettingsControlContext,
  field: ConfigSettingsField,
  mode: KeyInputMode,
): HTMLElement {
  const value = context.valueForField(field);
  return renderKeyLearnButton(typeof value === 'string' ? value : '', mode, (next) => {
    context.updateDraft(field.configPath, next);
  });
}

function applyMpvRows(
  context: SettingsControlContext,
  field: ConfigSettingsField,
  rows: MpvKeybindingRow[],
): void {
  context.updateDraft(field.configPath, buildMpvKeybindingConfigValue(DEFAULT_KEYBINDINGS, rows));
}

export function renderMpvKeybindingsInput(
  context: SettingsControlContext,
  field: ConfigSettingsField,
): HTMLElement {
  const rows = createMpvKeybindingRows(DEFAULT_KEYBINDINGS, context.valueForField(field));
  const container = createElement('div', 'keybinding-editor');

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i]!;
    const item = createElement('div', 'keybinding-row');
    const keyButton = renderKeyLearnButton(row.key, 'dom-code', (next) => {
      row.key = next;
      applyMpvRows(context, field, rows);
    });
    const command = createElement('input', 'config-input mono-input') as HTMLInputElement;
    command.type = 'text';
    command.value = row.commandText;
    command.placeholder = '["cycle","pause"]';
    command.addEventListener('input', () => {
      const parsed = parseMpvCommandText(command.value);
      if (parsed === undefined) {
        command.classList.add('invalid');
        context.setFieldError(field.configPath, 'Invalid MPV command JSON');
        return;
      }
      command.classList.remove('invalid');
      context.setFieldError(field.configPath, null);
      row.command = parsed;
      row.commandText = command.value;
      applyMpvRows(context, field, rows);
    });
    const removeButton = createElement('button', 'reset-button icon-button') as HTMLButtonElement;
    removeButton.type = 'button';
    removeButton.textContent = 'Remove';
    removeButton.addEventListener('click', () => {
      const nextRows = rows.filter((_, index) => index !== i);
      applyMpvRows(context, field, nextRows);
      requestRender();
    });
    item.append(keyButton, command, removeButton);
    container.append(item);
  }

  const addButton = createElement('button', 'secondary-button compact-button') as HTMLButtonElement;
  addButton.type = 'button';
  addButton.textContent = 'Add Binding';
  addButton.addEventListener('click', () => {
    rows.push({ defaultKey: '', key: '', command: null, commandText: '', isDefault: false });
    applyMpvRows(context, field, rows);
    requestRender();
  });
  container.append(addButton);

  return container;
}
