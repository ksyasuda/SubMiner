import type {
  ConfigSettingsAPI,
  ConfigSettingsCategory,
  ConfigSettingsField,
  ConfigSettingsPatchOperation,
  ConfigSettingsSnapshot,
  ConfigSettingsSnapshotValue,
} from '../types/settings';
import { i18n, applyI18nToDOM } from '../i18n/index.js';
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
import { getFieldTitleBadges } from './settings-field-layout';
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
  input: 'Input',
  integrations: 'Integrations',
  'tracking-app': 'Tracking & App',
  advanced: 'Advanced',
};

const CATEGORY_LABELS_I18N: Record<ConfigSettingsCategory, string> = {
  appearance: 'settingsCat.appearance',
  behavior: 'settingsCat.behavior',
  'mining-anki': 'settingsCat.mining',
  input: 'settingsCat.input',
  integrations: 'settingsCat.integrations',
  'tracking-app': 'settingsCat.tracking',
  advanced: 'settingsCat.advanced',
};

const SECTION_TITLE_I18N: Record<string, string> = {
  'Visible Overlay Auto-Start': 'settingsSection.visibleOverlayAutoStart',
  'UI Language': 'settingsSection.uiLanguage',
  'Texthooker Server': 'settingsSection.texthooker',
  'WebSocket Server': 'settingsSection.websocket',
  'Annotation WebSocket': 'settingsSection.annotationWebsocket',
  'Logging': 'settingsSection.logging',
  'Controller Support': 'settingsSection.controller',
  'Startup Warmups': 'settingsSection.startupWarmups',
  'Updates': 'settingsSection.updates',
  'Notifications': 'settingsSection.notifications',
  'Keyboard Shortcuts': 'settingsSection.keyboardShortcuts',
  'Keybindings (MPV Commands)': 'settingsSection.keybindings',
  'Secondary Subtitles': 'settingsSection.secondarySub',
  'Subtitle Sync': 'settingsSection.subsync',
  'Subtitle Position': 'settingsSection.subtitlePosition',
  'Subtitle Appearance': 'settingsSection.subtitleAppearance',
  'Subtitle Sidebar': 'settingsSection.subtitleSidebar',
  'Shared AI Provider': 'settingsSection.sharedAi',
  'AnkiConnect Integration': 'settingsSection.ankiConnect',
  'Jimaku': 'settingsSection.jimaku',
  'YouTube Playback Settings': 'settingsSection.youtube',
  'Anilist': 'settingsSection.anilist',
  'Yomitan': 'settingsSection.yomitan',
  'MPV Launcher': 'settingsSection.mpv',
  'Jellyfin': 'settingsSection.jellyfin',
  'Discord Rich Presence': 'settingsSection.discord',
  'Immersion Tracking': 'settingsSection.immersionTracking',
  'Stats Dashboard': 'settingsSection.stats',
};

function translateSectionTitle(title: string): string {
  const key = SECTION_TITLE_I18N[title];
  return key ? i18n.t(key) : title;
}

function fieldDescriptionKey(configPath: string): string {
  return `opt.${configPath}`;
}

function fieldLabelKey(configPath: string): string {
  return `optLabel.${configPath}`;
}

const CATEGORY_ORDER: ConfigSettingsCategory[] = [
  'appearance',
  'behavior',
  'mining-anki',
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
  dom.saveButton.textContent = dirtyCount > 0 ? i18n.t('settings.saveCount', { count: dirtyCount }) : i18n.t('settings.save');
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
  title.textContent = snapshot.warnings.length === 1
    ? i18n.t('settings.validationWarning', { count: snapshot.warnings.length })
    : i18n.t('settings.validationWarningPlural', { count: snapshot.warnings.length });
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
    label.textContent = i18n.t(CATEGORY_LABELS_I18N[category]);
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
  const labelText = createElement('span', 'field-title-text');
  labelText.textContent = i18n.t(fieldLabelKey(field.configPath), undefined, field.label);
  label.append(labelText);
  for (const badge of getFieldTitleBadges(field)) {
    const badgeEl = createElement('span', badge.className);
    badgeEl.textContent = badge.text;
    label.append(badgeEl);
  }
  const description = createElement('p');
  description.textContent = i18n.t(fieldDescriptionKey(field.configPath), undefined, field.description);
  header.append(label, description);

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
  resetButton.textContent = i18n.t('common.reset');
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
    dom.categoryTitle.textContent = i18n.t('settings.searchResults');
    dom.categoryMeta.textContent = i18n.t(fields.length === 1 ? 'settings.settingCount' : 'settings.settingCountPlural', { count: fields.length });
  } else {
    dom.categoryTitle.textContent = i18n.t(CATEGORY_LABELS_I18N[state.category]);
    dom.categoryMeta.textContent = i18n.t(fields.length === 1 ? 'settings.settingCount' : 'settings.settingCountPlural', { count: fields.length });
  }

  if (fields.length === 0) {
    const empty = createElement('div', 'empty-state');
    empty.textContent = i18n.t('settings.noMatch');
    dom.settingsContent.append(empty);
    return;
  }

  const sections = new Map<
    string,
    { title: string; rawSection: string; fields: ConfigSettingsField[] }
  >();
  for (const field of fields) {
    const title = query
      ? `${i18n.t(CATEGORY_LABELS_I18N[field.category])} / ${translateSectionTitle(field.section)}`
      : translateSectionTitle(field.section);
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
  setStatus(i18n.t('settings.saving'), 'info');
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

// Initialize i18n from main process (resolves OS language correctly) before loading settings
async function initI18n(): Promise<void> {
  try {
    const lang = await window.configSettingsAPI.getUILanguage();
    i18n.setLanguage(lang as 'en' | 'zh-CN');
  } catch {
    i18n.setLanguage('en');
  }
  applyI18nToDOM();
}

void initI18n().then(() => loadSnapshot()).catch((error) => {
  setStatus(error instanceof Error ? error.message : i18n.t('settings.loadFailed'), 'error');
});
