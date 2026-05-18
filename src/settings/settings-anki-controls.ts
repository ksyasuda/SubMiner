import type { ConfigSettingsField, ConfigSettingsSnapshotValue } from '../types/settings';
import type { SettingsControlContext } from './settings-control-context';
import { addOption, createElement, uniqueSorted } from './settings-control-dom';

const state: {
  deckNames: string[] | null;
  deckNamesLoading: boolean;
  deckNamesError: string | null;
  deckFieldNames: Map<string, string[]>;
  deckFieldNamesLoading: Set<string>;
  deckFieldNamesErrors: Map<string, string>;
  deckModelNames: Map<string, string[]>;
  deckModelNamesLoading: Map<string, Promise<string[]>>;
  modelNames: string[] | null;
  modelNamesLoading: boolean;
  modelNamesError: string | null;
  modelFieldNames: Map<string, string[]>;
  modelFieldNamesLoading: Set<string>;
  modelFieldNamesErrors: Map<string, string>;
  noteFieldModelName: string;
  ankiConnectUrl: string;
  noteFieldModelNameManuallySelected: boolean;
} = {
  deckNames: null,
  deckNamesLoading: false,
  deckNamesError: null,
  deckFieldNames: new Map(),
  deckFieldNamesLoading: new Set(),
  deckFieldNamesErrors: new Map(),
  deckModelNames: new Map(),
  deckModelNamesLoading: new Map(),
  modelNames: null,
  modelNamesLoading: false,
  modelNamesError: null,
  modelFieldNames: new Map(),
  modelFieldNamesLoading: new Set(),
  modelFieldNamesErrors: new Map(),
  noteFieldModelName: '',
  ankiConnectUrl: '',
  noteFieldModelNameManuallySelected: false,
};

let requestRender = (): void => undefined;

export function configureAnkiControls(options: { requestRender: () => void }): void {
  requestRender = options.requestRender;
}

export function initializeAnkiControls(_values: Record<string, ConfigSettingsSnapshotValue>): void {
  state.noteFieldModelName = '';
  state.noteFieldModelNameManuallySelected = false;
}

export function selectPreferredNoteFieldModelName(
  modelNames: readonly string[],
  currentModelName = '',
): string {
  void currentModelName;

  const exactKiku = modelNames.find((name) => name.trim().toLowerCase() === 'kiku');
  if (exactKiku) {
    return exactKiku;
  }

  const exactLapis = modelNames.find((name) => name.trim().toLowerCase() === 'lapis');
  if (exactLapis) {
    return exactLapis;
  }

  return '';
}

function normalizeStringArray(value: unknown): string[] {
  return Array.isArray(value)
    ? value.filter((entry): entry is string => typeof entry === 'string' && entry.length > 0)
    : [];
}

function normalizeKnownWordsDecks(value: unknown): Record<string, string[]> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return {};
  }
  const decks: Record<string, string[]> = {};
  for (const [deckName, fields] of Object.entries(value)) {
    if (!deckName) continue;
    decks[deckName] = normalizeStringArray(fields);
  }
  return decks;
}

function setKnownWordsDecks(
  context: SettingsControlContext,
  path: string,
  decks: Record<string, string[]>,
): void {
  const next: Record<string, string[]> = {};
  for (const [deckName, fields] of Object.entries(decks)) {
    if (!deckName) continue;
    next[deckName] = uniqueSorted(fields);
  }
  context.updateDraft(path, next);
}

function getDraftAnkiConnectUrl(context: SettingsControlContext): string | undefined {
  const value = context.valueForPath('ankiConnect.url');
  return typeof value === 'string' && value.trim().length > 0 ? value.trim() : undefined;
}

function getDraftAnkiDeckName(context: SettingsControlContext): string {
  const value = context.valueForPath('ankiConnect.deck');
  return typeof value === 'string' ? value.trim() : '';
}

function syncAnkiConnectUrl(draftUrl: string | undefined): void {
  const nextUrl = draftUrl ?? '';
  if (state.ankiConnectUrl === nextUrl) {
    return;
  }
  const hasAnkiMetadata =
    state.deckNames !== null ||
    state.deckNamesLoading ||
    state.deckFieldNames.size > 0 ||
    state.deckFieldNamesLoading.size > 0 ||
    state.deckModelNames.size > 0 ||
    state.deckModelNamesLoading.size > 0 ||
    state.modelNames !== null ||
    state.modelNamesLoading ||
    state.modelFieldNames.size > 0 ||
    state.modelFieldNamesLoading.size > 0;
  state.ankiConnectUrl = nextUrl;
  if (hasAnkiMetadata) {
    state.noteFieldModelName = '';
    state.noteFieldModelNameManuallySelected = false;
  }
  state.deckNames = null;
  state.deckNamesLoading = false;
  state.deckNamesError = null;
  state.deckFieldNames.clear();
  state.deckFieldNamesLoading.clear();
  state.deckFieldNamesErrors.clear();
  state.deckModelNames.clear();
  state.deckModelNamesLoading.clear();
  state.modelNames = null;
  state.modelNamesLoading = false;
  state.modelNamesError = null;
  state.modelFieldNames.clear();
  state.modelFieldNamesLoading.clear();
  state.modelFieldNamesErrors.clear();
}

async function loadAnkiDeckNames(draftUrl?: string): Promise<void> {
  syncAnkiConnectUrl(draftUrl);
  if (state.deckNames || state.deckNamesLoading) return;
  const requestUrl = state.ankiConnectUrl;
  state.deckNamesLoading = true;
  try {
    const result = await window.configSettingsAPI.getAnkiDeckNames(draftUrl);
    if (state.ankiConnectUrl !== requestUrl) return;
    if (result.ok) {
      state.deckNames = uniqueSorted(result.values);
      state.deckNamesError = null;
    } else {
      state.deckNames = [];
      state.deckNamesError = result.error ?? 'Failed to load Anki decks.';
    }
  } catch (error) {
    if (state.ankiConnectUrl !== requestUrl) return;
    state.deckNames = [];
    state.deckNamesError = error instanceof Error ? error.message : 'Failed to load Anki decks.';
  } finally {
    if (state.ankiConnectUrl === requestUrl) {
      state.deckNamesLoading = false;
      requestRender();
    }
  }
}

async function loadAnkiDeckFieldNames(deckName: string, draftUrl?: string): Promise<void> {
  syncAnkiConnectUrl(draftUrl);
  if (
    !deckName ||
    state.deckFieldNames.has(deckName) ||
    state.deckFieldNamesLoading.has(deckName)
  ) {
    return;
  }
  const requestUrl = state.ankiConnectUrl;
  state.deckFieldNamesLoading.add(deckName);
  try {
    const result = await window.configSettingsAPI.getAnkiDeckFieldNames(deckName, draftUrl);
    if (state.ankiConnectUrl !== requestUrl) return;
    if (result.ok) {
      state.deckFieldNames.set(deckName, uniqueSorted(result.values));
      state.deckFieldNamesErrors.delete(deckName);
    } else {
      state.deckFieldNames.set(deckName, []);
      state.deckFieldNamesErrors.set(
        deckName,
        result.error ?? `Failed to load fields for ${deckName}.`,
      );
    }
  } catch (error) {
    if (state.ankiConnectUrl !== requestUrl) return;
    state.deckFieldNames.set(deckName, []);
    state.deckFieldNamesErrors.set(
      deckName,
      error instanceof Error ? error.message : `Failed to load fields for ${deckName}.`,
    );
  } finally {
    if (state.ankiConnectUrl === requestUrl) {
      state.deckFieldNamesLoading.delete(deckName);
      requestRender();
    }
  }
}

async function loadAnkiDeckModelNames(deckName: string, draftUrl?: string): Promise<string[]> {
  syncAnkiConnectUrl(draftUrl);
  if (!deckName) return [];
  const cached = state.deckModelNames.get(deckName);
  if (cached) return cached;
  const loading = state.deckModelNamesLoading.get(deckName);
  if (loading) return loading;

  const requestUrl = state.ankiConnectUrl;
  const request = window.configSettingsAPI
    .getAnkiDeckModelNames(deckName, draftUrl)
    .then((result) => {
      if (state.ankiConnectUrl !== requestUrl) return [];
      const values = result.ok ? uniqueSorted(result.values) : [];
      state.deckModelNames.set(deckName, values);
      return values;
    })
    .catch(() => {
      if (state.ankiConnectUrl === requestUrl) {
        state.deckModelNames.set(deckName, []);
      }
      return [];
    })
    .finally(() => {
      if (state.ankiConnectUrl === requestUrl) {
        state.deckModelNamesLoading.delete(deckName);
        requestRender();
      }
    });

  state.deckModelNamesLoading.set(deckName, request);
  return request;
}

function findModelName(modelNames: readonly string[], modelName: string): string {
  const normalizedModelName = modelName.trim().toLowerCase();
  return normalizedModelName
    ? (modelNames.find((name) => name.toLowerCase() === normalizedModelName) ?? '')
    : '';
}

async function updatePreferredNoteFieldModelName(
  modelNames: readonly string[],
  deckName: string,
  draftUrl: string | undefined,
  requestUrl: string,
): Promise<void> {
  const deckModelNames = await loadAnkiDeckModelNames(deckName, draftUrl);
  if (state.ankiConnectUrl !== requestUrl || state.noteFieldModelNameManuallySelected) {
    return;
  }

  let nextModelName = '';
  for (const deckModelName of deckModelNames) {
    nextModelName = findModelName(modelNames, deckModelName);
    if (nextModelName) break;
  }
  nextModelName ||= selectPreferredNoteFieldModelName(modelNames, state.noteFieldModelName);

  if (state.noteFieldModelName !== nextModelName) {
    state.noteFieldModelName = nextModelName;
    requestRender();
  }
}

async function loadAnkiModelNames(draftUrl?: string, deckName = ''): Promise<void> {
  syncAnkiConnectUrl(draftUrl);
  if (state.modelNames) {
    if (!state.noteFieldModelNameManuallySelected) {
      void updatePreferredNoteFieldModelName(
        state.modelNames,
        deckName,
        draftUrl,
        state.ankiConnectUrl,
      );
    }
    return;
  }
  if (state.modelNamesLoading) return;
  const requestUrl = state.ankiConnectUrl;
  state.modelNamesLoading = true;
  try {
    const result = await window.configSettingsAPI.getAnkiModelNames(draftUrl);
    if (state.ankiConnectUrl !== requestUrl) return;
    if (result.ok) {
      state.modelNames = uniqueSorted(result.values);
      state.modelNamesError = null;
      if (!state.noteFieldModelNameManuallySelected) {
        await updatePreferredNoteFieldModelName(state.modelNames, deckName, draftUrl, requestUrl);
      }
    } else {
      state.modelNames = [];
      state.modelNamesError = result.error ?? 'Failed to load Anki note types.';
    }
  } catch (error) {
    if (state.ankiConnectUrl !== requestUrl) return;
    state.modelNames = [];
    state.modelNamesError =
      error instanceof Error ? error.message : 'Failed to load Anki note types.';
  } finally {
    if (state.ankiConnectUrl === requestUrl) {
      state.modelNamesLoading = false;
      requestRender();
    }
  }
}

async function loadAnkiModelFieldNames(modelName: string, draftUrl?: string): Promise<void> {
  syncAnkiConnectUrl(draftUrl);
  if (
    !modelName ||
    state.modelFieldNames.has(modelName) ||
    state.modelFieldNamesLoading.has(modelName)
  ) {
    return;
  }
  const requestUrl = state.ankiConnectUrl;
  state.modelFieldNamesLoading.add(modelName);
  try {
    const result = await window.configSettingsAPI.getAnkiModelFieldNames(modelName, draftUrl);
    if (state.ankiConnectUrl !== requestUrl) return;
    if (result.ok) {
      state.modelFieldNames.set(modelName, uniqueSorted(result.values));
      state.modelFieldNamesErrors.delete(modelName);
    } else {
      state.modelFieldNames.set(modelName, []);
      state.modelFieldNamesErrors.set(
        modelName,
        result.error ?? `Failed to load fields for ${modelName}.`,
      );
    }
  } catch (error) {
    if (state.ankiConnectUrl !== requestUrl) return;
    state.modelFieldNames.set(modelName, []);
    state.modelFieldNamesErrors.set(
      modelName,
      error instanceof Error ? error.message : `Failed to load fields for ${modelName}.`,
    );
  } finally {
    if (state.ankiConnectUrl === requestUrl) {
      state.modelFieldNamesLoading.delete(modelName);
      requestRender();
    }
  }
}

export function renderAnkiNoteTypeInput(
  context: SettingsControlContext,
  field: ConfigSettingsField,
): HTMLElement {
  const draftUrl = getDraftAnkiConnectUrl(context);
  void loadAnkiModelNames(draftUrl, getDraftAnkiDeckName(context));
  const currentValue = context.valueForField(field);
  const current = typeof currentValue === 'string' ? currentValue : '';
  const select = createElement('select', 'config-input') as HTMLSelectElement;
  const modelNames = uniqueSorted([...(state.modelNames ?? []), current]);
  if (state.modelNamesLoading && modelNames.length === 0) {
    addOption(select, current, 'Loading Note Types...');
  }
  for (const modelName of modelNames) {
    addOption(select, modelName);
  }
  select.value = current;
  select.addEventListener('change', () => context.updateDraft(field.configPath, select.value));

  const wrap = createElement('div', 'stacked-control');
  wrap.append(select);
  if (state.modelNamesError) {
    const hint = createElement('div', 'control-hint error');
    hint.textContent = state.modelNamesError;
    wrap.append(hint);
  }
  return wrap;
}

export function renderAnkiFieldInput(
  context: SettingsControlContext,
  field: ConfigSettingsField,
): HTMLElement {
  const draftUrl = getDraftAnkiConnectUrl(context);
  void loadAnkiModelNames(draftUrl, getDraftAnkiDeckName(context));
  if (state.noteFieldModelName) {
    void loadAnkiModelFieldNames(state.noteFieldModelName, draftUrl);
  }

  const currentValue = context.valueForField(field);
  const current = typeof currentValue === 'string' ? currentValue : '';
  const availableFields = state.noteFieldModelName
    ? (state.modelFieldNames.get(state.noteFieldModelName) ?? [])
    : [];
  const select = createElement('select', 'config-input') as HTMLSelectElement;
  if (!state.noteFieldModelName) {
    addOption(select, current, 'Select Note Type First');
    select.disabled = true;
  } else if (state.modelFieldNamesLoading.has(state.noteFieldModelName)) {
    addOption(select, current, current || 'Loading Fields...');
    select.disabled = true;
  } else {
    for (const fieldName of uniqueSorted([...availableFields, current])) {
      addOption(select, fieldName);
    }
  }
  select.value = current;
  select.addEventListener('change', () => context.updateDraft(field.configPath, select.value));

  const wrap = createElement('div', 'stacked-control');
  wrap.append(select);
  const error = state.modelFieldNamesErrors.get(state.noteFieldModelName);
  if (error) {
    const hint = createElement('div', 'control-hint error');
    hint.textContent = error;
    wrap.append(hint);
  }
  return wrap;
}

export function renderNoteFieldModelPicker(context: SettingsControlContext): HTMLElement {
  const draftUrl = getDraftAnkiConnectUrl(context);
  void loadAnkiModelNames(draftUrl, getDraftAnkiDeckName(context));
  if (state.noteFieldModelName) {
    void loadAnkiModelFieldNames(state.noteFieldModelName, draftUrl);
  }

  const row = createElement('article', 'field-row helper-row');
  const copy = createElement('div', 'field-copy');
  const title = createElement('h3');
  title.textContent = 'Note Type';
  const description = createElement('p');
  description.textContent =
    'Choose a note type from AnkiConnect to populate the field dropdowns below.';
  copy.append(title, description);

  const control = createElement('div', 'field-control');
  const select = createElement('select', 'config-input') as HTMLSelectElement;
  const modelNames = state.modelNames ?? [];
  if (state.modelNamesLoading && modelNames.length === 0) {
    addOption(select, '', 'Loading Note Types...');
  } else {
    for (const modelName of modelNames) {
      addOption(select, modelName);
    }
  }
  select.value = state.noteFieldModelName;
  select.addEventListener('change', () => {
    state.noteFieldModelName = select.value;
    state.noteFieldModelNameManuallySelected = true;
    requestRender();
  });
  control.append(select);
  row.append(copy, control);
  return row;
}

export function renderKnownWordsDecksInput(
  context: SettingsControlContext,
  field: ConfigSettingsField,
): HTMLElement {
  const draftUrl = getDraftAnkiConnectUrl(context);
  void loadAnkiDeckNames(draftUrl);
  const currentDecks = normalizeKnownWordsDecks(context.valueForField(field));
  const deckNames = state.deckNames ?? [];
  const container = createElement('div', 'deck-field-editor');

  const entries = Object.entries(currentDecks).sort(([left], [right]) => left.localeCompare(right));
  if (entries.length === 0) {
    const empty = createElement('div', 'control-hint');
    empty.textContent = state.deckNamesLoading
      ? 'Loading Anki decks...'
      : 'No known-word decks configured.';
    container.append(empty);
  }

  for (const [deckName, selectedFields] of entries) {
    if (deckName) {
      void loadAnkiDeckFieldNames(deckName, draftUrl);
    }
    const row = createElement('div', 'deck-field-row');
    const deckSelect = createElement('select', 'config-input') as HTMLSelectElement;
    for (const candidateDeck of uniqueSorted([...deckNames, deckName])) {
      addOption(deckSelect, candidateDeck);
    }
    deckSelect.value = deckName;
    deckSelect.addEventListener('change', () => {
      const nextDecks = normalizeKnownWordsDecks(context.valueForField(field));
      const fields = nextDecks[deckName] ?? [];
      delete nextDecks[deckName];
      nextDecks[deckSelect.value] = fields;
      setKnownWordsDecks(context, field.configPath, nextDecks);
      requestRender();
    });

    const availableFields = deckName ? (state.deckFieldNames.get(deckName) ?? []) : [];
    const fieldNames = uniqueSorted([...availableFields, ...selectedFields]);
    const fieldsWrap = createElement('div', 'deck-field-fields');
    const fieldActions = createElement('div', 'deck-field-actions');
    const checkboxList = createElement('div', 'field-checkbox-list');

    const setSelectedFields = (fields: string[]): void => {
      const nextDecks = normalizeKnownWordsDecks(context.valueForField(field));
      nextDecks[deckName] = fields;
      setKnownWordsDecks(context, field.configPath, nextDecks);
    };

    const selectAllButton = createElement(
      'button',
      'secondary-button compact-button',
    ) as HTMLButtonElement;
    selectAllButton.type = 'button';
    selectAllButton.textContent = 'Select All';
    selectAllButton.disabled = fieldNames.length === 0;
    selectAllButton.addEventListener('click', () => {
      setSelectedFields(fieldNames);
      requestRender();
    });

    const clearButton = createElement(
      'button',
      'secondary-button compact-button',
    ) as HTMLButtonElement;
    clearButton.type = 'button';
    clearButton.textContent = 'Clear';
    clearButton.disabled = selectedFields.length === 0;
    clearButton.addEventListener('click', () => {
      setSelectedFields([]);
      requestRender();
    });

    fieldActions.append(selectAllButton, clearButton);
    fieldsWrap.append(fieldActions, checkboxList);

    if (state.deckFieldNamesLoading.has(deckName)) {
      const hint = createElement('div', 'control-hint');
      hint.textContent = 'Loading Fields...';
      checkboxList.append(hint);
    } else if (fieldNames.length === 0) {
      const hint = createElement('div', 'control-hint');
      hint.textContent = deckName ? 'No fields found for this deck.' : 'Select A Deck First.';
      checkboxList.append(hint);
    }

    for (const candidateField of fieldNames) {
      const label = createElement('label', 'field-checkbox-row');
      const checkbox = createElement('input') as HTMLInputElement;
      checkbox.type = 'checkbox';
      checkbox.value = candidateField;
      checkbox.checked = selectedFields.includes(candidateField);
      checkbox.addEventListener('change', () => {
        const checkedFields = [
          ...checkboxList.querySelectorAll<HTMLInputElement>('input[type="checkbox"]:checked'),
        ].map((input) => input.value);
        setSelectedFields(checkedFields);
      });
      const text = createElement('span');
      text.textContent = candidateField;
      label.append(checkbox, text);
      checkboxList.append(label);
    }

    const removeButton = createElement('button', 'reset-button icon-button') as HTMLButtonElement;
    removeButton.type = 'button';
    removeButton.textContent = 'Remove';
    removeButton.addEventListener('click', () => {
      const nextDecks = normalizeKnownWordsDecks(context.valueForField(field));
      delete nextDecks[deckName];
      setKnownWordsDecks(context, field.configPath, nextDecks);
      requestRender();
    });

    row.append(deckSelect, fieldsWrap, removeButton);
    const error = state.deckFieldNamesErrors.get(deckName);
    if (error) {
      const hint = createElement('div', 'control-hint error');
      hint.textContent = error;
      row.append(hint);
    }
    container.append(row);
  }

  const addButton = createElement('button', 'secondary-button compact-button') as HTMLButtonElement;
  addButton.type = 'button';
  addButton.textContent = 'Add Deck';
  addButton.disabled = deckNames.length === 0;
  addButton.addEventListener('click', () => {
    const nextDecks = normalizeKnownWordsDecks(context.valueForField(field));
    const nextDeckName = deckNames.find((deckName) => !Object.hasOwn(nextDecks, deckName));
    if (!nextDeckName) return;
    nextDecks[nextDeckName] = [];
    setKnownWordsDecks(context, field.configPath, nextDecks);
    requestRender();
  });
  container.append(addButton);

  if (state.deckNamesError) {
    const hint = createElement('div', 'control-hint error');
    hint.textContent = state.deckNamesError;
    container.append(hint);
  }
  return container;
}
