import type { ConfigSettingsField, ConfigSettingsSnapshotValue } from '../types/settings';
import { toConfigDraftValue, toSettingsDisplayValue } from './settings-model';
import { parseOptionalNumberInputValue } from './input-values';
import {
  configureAnkiControls,
  renderAnkiDeckInput,
  initializeAnkiControls,
  renderAnkiFieldInput,
  renderAnkiNoteTypeInput,
  renderKnownWordsDecksInput,
  renderNoteFieldModelPicker,
} from './settings-anki-controls';
import type { SettingsControlContext } from './settings-control-context';
import { createElement, isSecretSnapshotValue } from './settings-control-dom';
import {
  configureKeybindingControls,
  renderKeyboardInput,
  renderMpvKeybindingsInput,
} from './settings-keybinding-controls';
import {
  getSubtitleCssManagedConfigPaths,
  getSubtitleCssScopeForPath,
  parseSubtitleCssDeclarations,
  serializeSubtitleCssDeclarations,
} from './subtitle-style-css';

export { renderNoteFieldModelPicker };

export function configureSettingsControls(options: { requestRender: () => void }): void {
  configureAnkiControls(options);
  configureKeybindingControls(options);
}

export function initializeSettingsControls(
  values: Record<string, ConfigSettingsSnapshotValue>,
): void {
  initializeAnkiControls(values);
}

function renderColorListInput(
  context: SettingsControlContext,
  field: ConfigSettingsField,
  value: ConfigSettingsSnapshotValue,
): HTMLElement {
  const colors = Array.isArray(value) ? (value as string[]) : [];
  const container = createElement('div', 'color-list');
  for (let i = 0; i < colors.length; i++) {
    const row = createElement('div', 'color-list-row');
    const label = createElement('span', 'color-list-label');
    label.textContent = `Band ${i + 1}`;
    const input = createElement('input', 'config-input') as HTMLInputElement;
    input.type = 'color';
    input.value = colors[i] ?? '#000000';
    input.addEventListener('input', () => {
      const updated = [...colors];
      updated[i] = input.value;
      context.updateDraft(field.configPath, updated);
    });
    row.append(label, input);
    container.append(row);
  }
  return container;
}

function renderJsonInput(
  context: SettingsControlContext,
  field: ConfigSettingsField,
  value: ConfigSettingsSnapshotValue,
): HTMLElement {
  const textarea = createElement('textarea', 'config-textarea') as HTMLTextAreaElement;
  textarea.spellcheck = false;
  textarea.value = JSON.stringify(value ?? {}, null, 2);
  textarea.addEventListener('input', () => {
    try {
      context.updateDraft(field.configPath, JSON.parse(textarea.value));
      textarea.classList.remove('invalid');
      context.setFieldError(field.configPath, null);
    } catch {
      textarea.classList.add('invalid');
      context.setFieldError(field.configPath, 'Invalid JSON');
    }
  });
  return textarea;
}

function renderStringListInput(
  context: SettingsControlContext,
  field: ConfigSettingsField,
  value: ConfigSettingsSnapshotValue,
): HTMLElement {
  const textarea = createElement('textarea', 'config-textarea compact') as HTMLTextAreaElement;
  textarea.spellcheck = false;
  textarea.value = Array.isArray(value) ? value.join('\n') : '';
  textarea.addEventListener('input', () => {
    context.updateDraft(
      field.configPath,
      textarea.value
        .split('\n')
        .map((entry) => entry.trim())
        .filter(Boolean),
    );
  });
  return textarea;
}

function renderCssDeclarationsInput(
  context: SettingsControlContext,
  field: ConfigSettingsField,
): HTMLElement {
  const scope = getSubtitleCssScopeForPath(field.configPath);
  const textarea = createElement(
    'textarea',
    'config-textarea css-declarations',
  ) as HTMLTextAreaElement;
  textarea.spellcheck = false;
  if (!scope) return textarea;

  const managedPaths = getSubtitleCssManagedConfigPaths(scope);
  const values: Record<string, ConfigSettingsSnapshotValue | undefined> = {
    [field.configPath]: context.valueForPath(field.configPath),
  };
  for (const path of managedPaths) {
    values[path] = context.valueForPath(path);
  }
  textarea.value = serializeSubtitleCssDeclarations(scope, values);
  textarea.addEventListener('input', () => {
    const parsed = parseSubtitleCssDeclarations(textarea.value);
    if (!parsed.ok) {
      textarea.classList.add('invalid');
      context.setFieldError(field.configPath, parsed.error);
      return;
    }

    textarea.classList.remove('invalid');
    context.setFieldError(field.configPath, null);
    context.updateDraft(field.configPath, parsed.declarations);
    for (const path of managedPaths) {
      context.resetDraftPath(path, undefined);
    }
  });
  return textarea;
}

export function renderControl(
  field: ConfigSettingsField,
  context: SettingsControlContext,
): HTMLElement {
  const value = toSettingsDisplayValue(field.configPath, context.valueForField(field));

  if (field.control === 'keyboard-shortcut') {
    return renderKeyboardInput(context, field, 'accelerator');
  }

  if (field.control === 'key-code') {
    return renderKeyboardInput(context, field, 'code');
  }

  if (field.control === 'mpv-key') {
    return renderKeyboardInput(context, field, 'mpv-key');
  }

  if (field.control === 'known-words-decks') {
    return renderKnownWordsDecksInput(context, field);
  }

  if (field.control === 'anki-deck') {
    return renderAnkiDeckInput(context, field);
  }

  if (field.control === 'anki-note-type') {
    return renderAnkiNoteTypeInput(context, field);
  }

  if (field.control === 'anki-field') {
    return renderAnkiFieldInput(context, field);
  }

  if (field.control === 'mpv-keybindings') {
    return renderMpvKeybindingsInput(context, field);
  }

  if (field.control === 'boolean') {
    const label = createElement('label', 'switch-control');
    const input = createElement('input') as HTMLInputElement;
    input.type = 'checkbox';
    input.checked = Boolean(value);
    input.addEventListener('change', () => context.updateDraft(field.configPath, input.checked));
    const track = createElement('span', 'switch-track');
    label.append(input, track);
    return label;
  }

  if (field.control === 'number') {
    const input = createElement('input', 'config-input') as HTMLInputElement;
    input.type = 'number';
    const numericValue =
      typeof value === 'number'
        ? value
        : typeof field.defaultValue === 'number'
          ? field.defaultValue
          : NaN;
    input.value = Number.isFinite(numericValue) ? String(numericValue) : '';
    input.addEventListener('input', () => {
      const next = parseOptionalNumberInputValue(input.value);
      if (next.ok) {
        input.classList.remove('invalid');
        context.setFieldError(field.configPath, null);
        context.updateDraft(field.configPath, toConfigDraftValue(field.configPath, next.value));
      } else {
        input.classList.add('invalid');
        context.setFieldError(field.configPath, 'Invalid number');
      }
    });
    return input;
  }

  if (field.control === 'select') {
    const select = createElement('select', 'config-input') as HTMLSelectElement;
    const enumValues = field.enumValues ?? [];
    if (typeof value === 'string' && value.length > 0 && !enumValues.includes(value)) {
      const option = createElement('option') as HTMLOptionElement;
      option.value = value;
      option.textContent = `${value} (config file only)`;
      option.selected = true;
      select.append(option);
    }
    for (const enumValue of enumValues) {
      const option = createElement('option') as HTMLOptionElement;
      option.value = enumValue;
      option.textContent = enumValue;
      option.selected = enumValue === value;
      select.append(option);
    }
    select.addEventListener('change', () => context.updateDraft(field.configPath, select.value));
    return select;
  }

  if (field.control === 'color-list') {
    return renderColorListInput(context, field, value);
  }

  if (field.control === 'string-list') {
    return renderStringListInput(context, field, value);
  }

  if (field.control === 'json') {
    return renderJsonInput(context, field, value);
  }

  if (field.control === 'css-declarations') {
    return renderCssDeclarationsInput(context, field);
  }

  if (field.control === 'textarea') {
    const textarea = createElement('textarea', 'config-textarea compact') as HTMLTextAreaElement;
    textarea.spellcheck = false;
    textarea.value = typeof value === 'string' ? value : '';
    textarea.addEventListener('input', () => context.updateDraft(field.configPath, textarea.value));
    return textarea;
  }

  const input = createElement('input', 'config-input') as HTMLInputElement;
  input.type = field.control === 'secret' ? 'password' : field.control;
  if (field.control === 'secret') {
    input.placeholder =
      isSecretSnapshotValue(value) && value.configured ? 'Configured' : 'Not configured';
    input.addEventListener('input', () => {
      if (input.value.trim().length === 0) {
        context.updateDraft(field.configPath, value);
        return;
      }
      context.updateDraft(field.configPath, input.value);
    });
  } else {
    input.value = typeof value === 'string' ? value : '';
    input.addEventListener('input', () => context.updateDraft(field.configPath, input.value));
  }
  return input;
}
