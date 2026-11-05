import { getSubtitleFilterPatternError } from '../shared/subtitle-text-filter';
import type { ConfigSettingsField } from '../types/settings';
import type { SubtitleFilterRule } from '../types/subtitle';
import type { SettingsControlContext } from './settings-control-context';
import { createElement } from './settings-control-dom';

let requestRender = (): void => undefined;

export function configureSubtitleFilterControls(options: { requestRender: () => void }): void {
  requestRender = options.requestRender;
}

function toRules(value: unknown): SubtitleFilterRule[] {
  if (!Array.isArray(value)) return [];
  return value
    .filter(
      (rule): rule is Partial<SubtitleFilterRule> => rule !== null && typeof rule === 'object',
    )
    .map((rule) => ({
      name: rule.name ?? '',
      pattern: rule.pattern ?? '',
      replacement: rule.replacement ?? '',
      enabled: rule.enabled ?? true,
    }));
}

function renderTextInput(
  value: string,
  placeholder: string,
  className: string,
  onInput: (value: string) => void,
): HTMLInputElement {
  const input = createElement('input', className) as HTMLInputElement;
  input.type = 'text';
  input.spellcheck = false;
  input.value = value;
  input.placeholder = placeholder;
  input.addEventListener('input', () => onInput(input.value));
  return input;
}

/** Editor for `subtitleFilter.customRules`: one card per rule, validated as you type. */
export function renderSubtitleFilterRulesInput(
  context: SettingsControlContext,
  field: ConfigSettingsField,
): HTMLElement {
  const rules = toRules(context.valueForField(field));
  const patternErrors = new Map<number, string>();
  const container = createElement('div', 'stacked-control');

  // An invalid pattern blocks saving; empty patterns are inert placeholders.
  const syncFieldError = (): void => {
    const invalid = [...patternErrors.keys()].sort((a, b) => a - b);
    context.setFieldError(
      field.configPath,
      invalid.length > 0 ? `Invalid pattern in rule ${invalid[0]! + 1}` : null,
    );
  };
  const commit = (): void => {
    syncFieldError();
    context.updateDraft(
      field.configPath,
      rules.map((rule) => ({ ...rule })),
    );
  };

  rules.forEach((rule, index) => {
    const card = createElement('div', 'deck-field-row');
    const header = createElement('div', 'deck-field-row-header');

    const toggle = createElement('label', 'switch-control');
    toggle.title = 'Enabled';
    const checkbox = createElement('input') as HTMLInputElement;
    checkbox.type = 'checkbox';
    checkbox.checked = rule.enabled;
    checkbox.addEventListener('change', () => {
      rule.enabled = checkbox.checked;
      commit();
    });
    toggle.append(checkbox, createElement('span', 'switch-track'));

    const name = renderTextInput(
      rule.name,
      'Rule name',
      'config-input deck-field-row-name',
      (v) => {
        rule.name = v;
        commit();
      },
    );

    const removeButton = createElement('button', 'reset-button icon-button') as HTMLButtonElement;
    removeButton.type = 'button';
    removeButton.textContent = 'Remove';
    removeButton.addEventListener('click', () => {
      rules.splice(index, 1);
      commit();
      requestRender();
    });
    header.append(toggle, name, removeButton);

    const errorHint = createElement('div', 'control-hint error');
    const showPatternError = (input: HTMLInputElement): void => {
      const error = input.value ? getSubtitleFilterPatternError(input.value) : null;
      input.classList.toggle('invalid', error !== null);
      errorHint.textContent = error ?? '';
      errorHint.hidden = error === null;
      if (error === null) patternErrors.delete(index);
      else patternErrors.set(index, error);
    };
    const pattern = renderTextInput(
      rule.pattern,
      'Regex pattern, e.g. [（(][^）)]*[）)]',
      'config-input mono-input',
      (v) => {
        rule.pattern = v;
        showPatternError(pattern);
        commit();
      },
    );
    showPatternError(pattern);

    const replacement = renderTextInput(
      rule.replacement,
      'Replacement (leave empty to remove the match)',
      'config-input mono-input',
      (v) => {
        rule.replacement = v;
        commit();
      },
    );

    card.append(header, pattern, errorHint, replacement);
    container.append(card);
  });

  syncFieldError();

  const addButton = createElement('button', 'secondary-button compact-button') as HTMLButtonElement;
  addButton.type = 'button';
  addButton.textContent = 'Add Rule';
  addButton.addEventListener('click', () => {
    rules.push({ name: '', pattern: '', replacement: '', enabled: true });
    context.updateDraft(
      field.configPath,
      rules.map((rule) => ({ ...rule })),
    );
    requestRender();
  });
  container.append(addButton);

  return container;
}
