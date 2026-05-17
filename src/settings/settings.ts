import type {
  ConfigSettingsAPI,
  ConfigSettingsCategory,
  ConfigSettingsField,
  ConfigSettingsPatchOperation,
  ConfigSettingsSnapshot,
  ConfigSettingsSnapshotValue,
} from '../types/settings';
import {
  configureSettingsControls,
  initializeSettingsControls,
  renderControl,
  renderNoteFieldModelPicker,
} from './settings-controls';
import {
  createSettingsDraft,
  filterSettingsFields,
  getDirtyOperations,
  resetDraftPath,
  setDraftValue,
  type SettingsDraft,
} from './settings-model';
import { getSubtitleCssManagedConfigPaths, getSubtitleCssScopeForPath } from './subtitle-style-css';

declare global {
  interface Window {
    configSettingsAPI: ConfigSettingsAPI;
  }
}

const CATEGORY_LABELS: Record<ConfigSettingsCategory, string> = {
  appearance: 'Appearance',
  behavior: 'Behavior',
  'mining-anki': 'Mining & Anki',
  'playback-sources': 'Playback & Sources',
  input: 'Input',
  integrations: 'Integrations',
  'tracking-app': 'Tracking & App',
  advanced: 'Advanced',
};

const CATEGORY_ORDER: ConfigSettingsCategory[] = [
  'appearance',
  'behavior',
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
  category: 'appearance',
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
  saveButton: getElement<HTMLButtonElement>('saveButton'),
  statusBanner: getElement<HTMLElement>('statusBanner'),
  warningsPanel: getElement<HTMLElement>('warningsPanel'),
  settingsContent: getElement<HTMLElement>('settingsContent'),
};

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
  if (!state.draft) {
    return field.defaultValue;
  }
  return Object.hasOwn(state.draft.values, field.configPath)
    ? state.draft.values[field.configPath]
    : field.defaultValue;
}

function valueForPath(path: string): ConfigSettingsSnapshotValue | undefined {
  if (!state.draft || !Object.hasOwn(state.draft.values, path)) {
    return undefined;
  }
  return state.draft.values[path];
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

function resetDraftPathContext(path: string, defaultValue?: ConfigSettingsSnapshotValue): void {
  if (!state.draft) return;
  resetDraftPath(state.draft, path, defaultValue);
  state.inputErrors.delete(path);
  syncSaveButton();
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
      (field) => field.category === category && !field.legacyHidden && !field.settingsHidden,
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
      dom.settingsContent.scrollTop = 0;
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
  controlWrap.append(
    renderControl(field, {
      setFieldError,
      resetDraftPath: resetDraftPathContext,
      updateDraft,
      valueForField,
      valueForPath,
    }),
  );
  const resetButton = createElement('button', 'reset-button') as HTMLButtonElement;
  resetButton.type = 'button';
  resetButton.textContent = 'Reset';
  resetButton.addEventListener('click', () => {
    if (!state.draft) return;
    resetDraftPath(state.draft, field.configPath, field.defaultValue);
    const cssScope = getSubtitleCssScopeForPath(field.configPath);
    if (cssScope) {
      for (const path of getSubtitleCssManagedConfigPaths(cssScope)) {
        resetDraftPath(state.draft, path, undefined);
      }
    }
    state.inputErrors.delete(field.configPath);
    render();
  });
  controlWrap.append(resetButton);
  row.append(header, controlWrap);
  return row;
}

function renderSettingsContent(snapshot: ConfigSettingsSnapshot): void {
  dom.settingsContent.replaceChildren();
  const query = state.query.trim();
  const fields = filterSettingsFields(snapshot.fields, {
    category: query ? undefined : state.category,
    query,
  });

  if (query) {
    const categoryCount = new Set(fields.map((field) => field.category)).size;
    dom.categoryTitle.textContent = 'Search results';
    dom.categoryMeta.textContent = `${fields.length} setting${fields.length === 1 ? '' : 's'}${
      categoryCount > 0
        ? ` across ${categoryCount} categor${categoryCount === 1 ? 'y' : 'ies'}`
        : ''
    }`;
  } else {
    dom.categoryTitle.textContent = CATEGORY_LABELS[state.category];
    dom.categoryMeta.textContent = `${fields.length} setting${fields.length === 1 ? '' : 's'}`;
  }

  if (fields.length === 0) {
    const empty = createElement('div', 'empty-state');
    empty.textContent = 'No matching settings';
    dom.settingsContent.append(empty);
    return;
  }

  const sections = new Map<
    string,
    { title: string; rawSection: string; fields: ConfigSettingsField[] }
  >();
  for (const field of fields) {
    const title = query ? `${CATEGORY_LABELS[field.category]} / ${field.section}` : field.section;
    const section = sections.get(title) ?? { title, rawSection: field.section, fields: [] };
    section.fields.push(field);
    sections.set(title, section);
  }

  for (const section of sections.values()) {
    const sectionEl = createElement('section', 'settings-section');
    const title = createElement('h2');
    title.textContent = section.title;
    sectionEl.append(title);
    if (section.rawSection === 'Note Fields') {
      sectionEl.append(
        renderNoteFieldModelPicker({
          setFieldError,
          resetDraftPath: resetDraftPathContext,
          updateDraft,
          valueForField,
          valueForPath,
        }),
      );
    }
    let currentSubsection = '';
    for (const field of section.fields) {
      if (field.subsection && field.subsection !== currentSubsection) {
        currentSubsection = field.subsection;
        const subsectionTitle = createElement('h3', 'settings-subsection-title');
        subsectionTitle.textContent = field.subsection;
        sectionEl.append(subsectionTitle);
      }
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

configureSettingsControls({ requestRender: render });

async function loadSnapshot(): Promise<void> {
  clearStatus();
  const snapshot = await window.configSettingsAPI.getSnapshot();
  state.snapshot = snapshot;
  state.draft = createSettingsDraft(snapshot.values);
  initializeSettingsControls(snapshot.values);
  state.inputErrors.clear();
  render();
}

async function save(): Promise<void> {
  if (!state.draft) return;
  const operations: ConfigSettingsPatchOperation[] = getDirtyOperations(state.draft);
  if (operations.length === 0) return;

  dom.saveButton.disabled = true;
  setStatus('Saving...', 'info');
  let result;
  try {
    result = await window.configSettingsAPI.savePatch({ operations });
  } catch (error) {
    setStatus(error instanceof Error ? error.message : 'Save failed', 'error');
    syncSaveButton();
    return;
  }
  if (!result.ok || !result.snapshot) {
    const message =
      result.error ??
      result.warnings?.map((warning) => `${warning.path}: ${warning.message}`).join('\n') ??
      'Save failed';
    setStatus(message, 'error');
    syncSaveButton();
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
}

dom.searchInput.addEventListener('input', () => {
  state.query = dom.searchInput.value;
  render();
});
dom.saveButton.addEventListener('click', () => {
  void save();
});

void loadSnapshot().catch((error) => {
  setStatus(error instanceof Error ? error.message : 'Failed to load settings', 'error');
});
