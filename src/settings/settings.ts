import type {
  ConfigSettingsAPI,
  ConfigSettingsCategory,
  ConfigSettingsField,
  ConfigSettingsPatchOperation,
  ConfigSettingsSnapshot,
  ConfigSettingsSnapshotValue,
} from '../types/settings';
import { parseOptionalNumberInputValue } from './input-values';
import {
  createSettingsDraft,
  filterSettingsFields,
  getDirtyOperations,
  resetDraftPath,
  setDraftValue,
  type SettingsDraft,
} from './settings-model';

declare global {
  interface Window {
    configSettingsAPI: ConfigSettingsAPI;
  }
}

const CATEGORY_LABELS: Record<ConfigSettingsCategory, string> = {
  viewing: 'Viewing',
  'mining-anki': 'Mining & Anki',
  'playback-sources': 'Playback & Sources',
  input: 'Input',
  integrations: 'Integrations',
  'tracking-app': 'Tracking & App',
  advanced: 'Advanced',
};

const CATEGORY_ORDER: ConfigSettingsCategory[] = [
  'viewing',
  'mining-anki',
  'playback-sources',
  'input',
  'integrations',
  'tracking-app',
  'advanced',
];

const state: {
  snapshot: ConfigSettingsSnapshot | null;
  draft: SettingsDraft | null;
  category: ConfigSettingsCategory;
  query: string;
  inputErrors: Map<string, string>;
} = {
  snapshot: null,
  draft: null,
  category: 'viewing',
  query: '',
  inputErrors: new Map(),
};

function getElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Missing settings element: ${id}`);
  }
  return element as T;
}

const dom = {
  categoryNav: getElement<HTMLElement>('categoryNav'),
  categoryTitle: getElement<HTMLHeadingElement>('categoryTitle'),
  categoryMeta: getElement<HTMLElement>('categoryMeta'),
  searchInput: getElement<HTMLInputElement>('searchInput'),
  openFileButton: getElement<HTMLButtonElement>('openFileButton'),
  saveButton: getElement<HTMLButtonElement>('saveButton'),
  statusBanner: getElement<HTMLElement>('statusBanner'),
  warningsPanel: getElement<HTMLElement>('warningsPanel'),
  settingsContent: getElement<HTMLElement>('settingsContent'),
};

function isSecretSnapshotValue(
  value: ConfigSettingsSnapshotValue,
): value is { configured: boolean } {
  return Boolean(value && typeof value === 'object' && 'configured' in value);
}

function setStatus(message: string, tone: 'info' | 'error' | 'success' = 'info'): void {
  dom.statusBanner.textContent = message;
  dom.statusBanner.className = `status-banner ${tone}`;
}

function clearStatus(): void {
  dom.statusBanner.textContent = '';
  dom.statusBanner.className = 'status-banner hidden';
}

function getDirtyCount(): number {
  return state.draft ? getDirtyOperations(state.draft).length : 0;
}

function syncSaveButton(): void {
  const dirtyCount = getDirtyCount();
  dom.saveButton.disabled = dirtyCount === 0 || state.inputErrors.size > 0;
  dom.saveButton.textContent = dirtyCount > 0 ? `Save ${dirtyCount}` : 'Save';
}

function createElement<K extends keyof HTMLElementTagNameMap>(
  tagName: K,
  className?: string,
): HTMLElementTagNameMap[K] {
  const element = document.createElement(tagName);
  if (className) {
    element.className = className;
  }
  return element;
}

function createFieldMeta(field: ConfigSettingsField): HTMLElement {
  const meta = createElement('div', 'field-meta');
  const path = createElement('code');
  path.textContent = field.configPath;
  meta.append(path);

  const restart = createElement('span', `restart-chip ${field.restartBehavior}`);
  restart.textContent = field.restartBehavior === 'hot-reload' ? 'Live' : 'Restart';
  meta.append(restart);

  if (field.advanced) {
    const advanced = createElement('span', 'advanced-chip');
    advanced.textContent = 'Advanced';
    meta.append(advanced);
  }
  return meta;
}

function valueForField(field: ConfigSettingsField): ConfigSettingsSnapshotValue {
  return state.draft?.values[field.configPath] ?? field.defaultValue;
}

function setFieldError(path: string, message: string | null): void {
  if (message) {
    state.inputErrors.set(path, message);
  } else {
    state.inputErrors.delete(path);
  }
  syncSaveButton();
}

function updateDraft(path: string, value: ConfigSettingsSnapshotValue): void {
  if (!state.draft) return;
  setDraftValue(state.draft, path, value);
  syncSaveButton();
}

function renderJsonInput(
  field: ConfigSettingsField,
  value: ConfigSettingsSnapshotValue,
): HTMLElement {
  const textarea = createElement('textarea', 'config-textarea') as HTMLTextAreaElement;
  textarea.spellcheck = false;
  textarea.value = JSON.stringify(value ?? {}, null, 2);
  textarea.addEventListener('input', () => {
    try {
      updateDraft(field.configPath, JSON.parse(textarea.value));
      textarea.classList.remove('invalid');
      setFieldError(field.configPath, null);
    } catch {
      textarea.classList.add('invalid');
      setFieldError(field.configPath, 'Invalid JSON');
    }
  });
  return textarea;
}

function renderStringListInput(
  field: ConfigSettingsField,
  value: ConfigSettingsSnapshotValue,
): HTMLElement {
  const textarea = createElement('textarea', 'config-textarea compact') as HTMLTextAreaElement;
  textarea.spellcheck = false;
  textarea.value = Array.isArray(value) ? value.join('\n') : '';
  textarea.addEventListener('input', () => {
    updateDraft(
      field.configPath,
      textarea.value
        .split('\n')
        .map((entry) => entry.trim())
        .filter(Boolean),
    );
  });
  return textarea;
}

function renderControl(field: ConfigSettingsField): HTMLElement {
  const value = valueForField(field);

  if (field.control === 'boolean') {
    const label = createElement('label', 'switch-control');
    const input = createElement('input') as HTMLInputElement;
    input.type = 'checkbox';
    input.checked = Boolean(value);
    input.addEventListener('change', () => updateDraft(field.configPath, input.checked));
    const track = createElement('span', 'switch-track');
    label.append(input, track);
    return label;
  }

  if (field.control === 'number') {
    const input = createElement('input', 'config-input') as HTMLInputElement;
    input.type = 'number';
    input.value = typeof value === 'number' ? String(value) : '';
    input.addEventListener('input', () => {
      const next = parseOptionalNumberInputValue(input.value);
      if (next.ok) {
        input.classList.remove('invalid');
        setFieldError(field.configPath, null);
        updateDraft(field.configPath, next.value);
      } else {
        input.classList.add('invalid');
        setFieldError(field.configPath, 'Invalid number');
      }
    });
    return input;
  }

  if (field.control === 'select') {
    const select = createElement('select', 'config-input') as HTMLSelectElement;
    for (const enumValue of field.enumValues ?? []) {
      const option = createElement('option') as HTMLOptionElement;
      option.value = enumValue;
      option.textContent = enumValue;
      option.selected = enumValue === value;
      select.append(option);
    }
    select.addEventListener('change', () => updateDraft(field.configPath, select.value));
    return select;
  }

  if (field.control === 'string-list') {
    return renderStringListInput(field, value);
  }

  if (field.control === 'json') {
    return renderJsonInput(field, value);
  }

  if (field.control === 'textarea') {
    const textarea = createElement('textarea', 'config-textarea compact') as HTMLTextAreaElement;
    textarea.spellcheck = false;
    textarea.value = typeof value === 'string' ? value : '';
    textarea.addEventListener('input', () => updateDraft(field.configPath, textarea.value));
    return textarea;
  }

  const input = createElement('input', 'config-input') as HTMLInputElement;
  input.type = field.control === 'secret' ? 'password' : field.control;
  if (field.control === 'secret') {
    input.placeholder =
      isSecretSnapshotValue(value) && value.configured ? 'Configured' : 'Not configured';
    input.addEventListener('input', () => {
      if (input.value.trim().length === 0) {
        if (state.draft) {
          setDraftValue(state.draft, field.configPath, state.draft.initialValues[field.configPath]);
        }
        syncSaveButton();
        return;
      }
      updateDraft(field.configPath, input.value);
    });
  } else {
    input.value = typeof value === 'string' ? value : '';
    input.addEventListener('input', () => updateDraft(field.configPath, input.value));
  }
  return input;
}

function renderWarnings(snapshot: ConfigSettingsSnapshot): void {
  dom.warningsPanel.replaceChildren();
  if (snapshot.warnings.length === 0) {
    dom.warningsPanel.className = 'warnings-panel hidden';
    return;
  }

  const title = createElement('div', 'warnings-title');
  title.textContent = `${snapshot.warnings.length} validation warning${
    snapshot.warnings.length === 1 ? '' : 's'
  }`;
  dom.warningsPanel.append(title);

  for (const warning of snapshot.warnings.slice(0, 6)) {
    const row = createElement('div', 'warning-row');
    const path = createElement('code');
    path.textContent = warning.path;
    const message = createElement('span');
    message.textContent = warning.message;
    row.append(path, message);
    dom.warningsPanel.append(row);
  }
  dom.warningsPanel.className = 'warnings-panel';
}

function renderCategoryNav(snapshot: ConfigSettingsSnapshot): void {
  dom.categoryNav.replaceChildren();
  for (const category of CATEGORY_ORDER) {
    const count = snapshot.fields.filter(
      (field) => field.category === category && !field.legacyHidden,
    ).length;
    if (count === 0) continue;
    const button = createElement('button', 'category-button') as HTMLButtonElement;
    button.type = 'button';
    button.classList.toggle('active', state.category === category);
    const label = createElement('span');
    label.textContent = CATEGORY_LABELS[category];
    const badge = createElement('strong');
    badge.textContent = String(count);
    button.append(label, badge);
    button.addEventListener('click', () => {
      state.category = category;
      render();
    });
    dom.categoryNav.append(button);
  }
}

function renderField(field: ConfigSettingsField): HTMLElement {
  const row = createElement('article', 'field-row');
  const header = createElement('div', 'field-copy');
  const label = createElement('h3');
  label.textContent = field.label;
  const description = createElement('p');
  description.textContent = field.description;
  header.append(label, description, createFieldMeta(field));

  const controlWrap = createElement('div', 'field-control');
  controlWrap.append(renderControl(field));
  const resetButton = createElement('button', 'reset-button') as HTMLButtonElement;
  resetButton.type = 'button';
  resetButton.textContent = 'Reset';
  resetButton.addEventListener('click', () => {
    if (!state.draft) return;
    resetDraftPath(state.draft, field.configPath, field.defaultValue);
    state.inputErrors.delete(field.configPath);
    render();
  });
  controlWrap.append(resetButton);
  row.append(header, controlWrap);
  return row;
}

function renderSettingsContent(snapshot: ConfigSettingsSnapshot): void {
  dom.settingsContent.replaceChildren();
  const fields = filterSettingsFields(snapshot.fields, {
    category: state.category,
    query: state.query,
  });

  dom.categoryTitle.textContent = CATEGORY_LABELS[state.category];
  dom.categoryMeta.textContent = `${fields.length} setting${fields.length === 1 ? '' : 's'}`;

  if (fields.length === 0) {
    const empty = createElement('div', 'empty-state');
    empty.textContent = 'No matching settings';
    dom.settingsContent.append(empty);
    return;
  }

  const sections = new Map<string, ConfigSettingsField[]>();
  for (const field of fields) {
    const sectionFields = sections.get(field.section) ?? [];
    sectionFields.push(field);
    sections.set(field.section, sectionFields);
  }

  for (const [section, sectionFields] of sections) {
    const sectionEl = createElement('section', 'settings-section');
    const title = createElement('h2');
    title.textContent = section;
    sectionEl.append(title);
    for (const field of sectionFields) {
      sectionEl.append(renderField(field));
    }
    dom.settingsContent.append(sectionEl);
  }
}

function render(): void {
  const snapshot = state.snapshot;
  if (!snapshot) return;
  renderCategoryNav(snapshot);
  renderWarnings(snapshot);
  renderSettingsContent(snapshot);
  syncSaveButton();
}

async function loadSnapshot(): Promise<void> {
  clearStatus();
  const snapshot = await window.configSettingsAPI.getSnapshot();
  state.snapshot = snapshot;
  state.draft = createSettingsDraft(snapshot.values);
  state.inputErrors.clear();
  render();
}

async function save(): Promise<void> {
  if (!state.draft) return;
  const operations: ConfigSettingsPatchOperation[] = getDirtyOperations(state.draft);
  if (operations.length === 0) return;

  dom.saveButton.disabled = true;
  setStatus('Saving...', 'info');
  try {
    const result = await window.configSettingsAPI.savePatch({ operations });
    if (!result.ok || !result.snapshot) {
      const message =
        result.error ??
        result.warnings?.map((warning) => `${warning.path}: ${warning.message}`).join('\n') ??
        'Save failed';
      setStatus(message, 'error');
      return;
    }

    state.snapshot = result.snapshot;
    state.draft = createSettingsDraft(result.snapshot.values);
    state.inputErrors.clear();
    const restartSections = result.restartRequiredSections ?? [];
    if (restartSections.length > 0) {
      setStatus(`Saved. Restart required: ${restartSections.join(', ')}`, 'info');
    } else if (result.hotReloadFields.length > 0) {
      setStatus('Saved. Live settings applied.', 'success');
    } else {
      setStatus('Saved.', 'success');
    }
    render();
  } catch (error) {
    setStatus(error instanceof Error ? error.message : 'Save failed', 'error');
  } finally {
    syncSaveButton();
  }
}

dom.searchInput.addEventListener('input', () => {
  state.query = dom.searchInput.value;
  render();
});
dom.saveButton.addEventListener('click', () => {
  void save();
});
dom.openFileButton.addEventListener('click', () => {
  void window.configSettingsAPI.openSettingsFile();
});

void loadSnapshot().catch((error) => {
  setStatus(error instanceof Error ? error.message : 'Failed to load settings', 'error');
});
