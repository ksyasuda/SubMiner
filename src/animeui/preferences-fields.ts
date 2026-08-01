import { describe } from './dom';
import type { SourcePreferenceView } from '../types/anime-browser';

/**
 * Renders one extension's settings schema as form fields.
 *
 * The extension owns the schema, so a commit hands back a refreshed one and the
 * whole panel re-renders from it — the Jellyfin source fills in its library
 * picker only after a successful login, and that has to show up.
 */

export type PreferenceCommit = (
  key: string,
  value: string | string[] | boolean,
) => Promise<SourcePreferenceView[]>;

/** Masked in the UI so a shoulder-surfer cannot read a stored password. */
function isSecretKey(view: SourcePreferenceView): boolean {
  return /password|token|api[-_ ]?key|secret/i.test(`${view.key} ${view.title}`);
}

function renderPreferenceField(
  container: HTMLElement,
  view: SourcePreferenceView,
  commit: PreferenceCommit,
): HTMLElement {
  const field = document.createElement('div');
  field.className = 'field';

  const label = document.createElement('label');
  label.className = 'field-label';
  label.textContent = view.title;
  field.append(label);

  if (view.summary) {
    const summary = document.createElement('span');
    summary.className = 'field-summary';
    summary.textContent = view.summary;
    field.append(summary);
  }

  const state = document.createElement('span');
  state.className = 'field-state';

  const save = async (value: string | string[] | boolean): Promise<void> => {
    state.removeAttribute('data-tone');
    state.textContent = 'Saving…';
    try {
      const refreshed = await commit(view.key, value);
      state.dataset.tone = 'ok';
      state.textContent = 'Saved';
      renderPreferences(container, refreshed, commit);
    } catch (error) {
      state.dataset.tone = 'error';
      state.textContent = describe(error);
    }
  };

  const row = document.createElement('div');
  row.className = 'field-row';

  if (view.kind === 'switch') {
    const wrapper = document.createElement('label');
    wrapper.className = 'field-check';
    const box = document.createElement('input');
    box.type = 'checkbox';
    box.checked = view.value === true;
    box.addEventListener('change', () => void save(box.checked));
    wrapper.append(box, document.createTextNode('Enabled'));
    row.append(wrapper);
  } else if (view.kind === 'list') {
    const select = document.createElement('select');
    select.className = 'text-input';
    for (const [index, entryValue] of view.entryValues.entries()) {
      const option = document.createElement('option');
      option.value = entryValue;
      option.textContent = view.entries[index] ?? entryValue;
      option.selected = entryValue === view.value;
      select.append(option);
    }
    if (view.entryValues.length === 0) {
      select.disabled = true;
      const option = document.createElement('option');
      option.textContent = 'Nothing to choose yet';
      select.append(option);
    }
    select.addEventListener('change', () => void save(select.value));
    row.append(select);
  } else if (view.kind === 'multi') {
    const group = document.createElement('div');
    group.className = 'field-multi';
    const selected = new Set(Array.isArray(view.value) ? view.value : []);
    for (const [index, entryValue] of view.entryValues.entries()) {
      const wrapper = document.createElement('label');
      wrapper.className = 'field-check';
      const box = document.createElement('input');
      box.type = 'checkbox';
      box.checked = selected.has(entryValue);
      box.addEventListener('change', () => {
        if (box.checked) selected.add(entryValue);
        else selected.delete(entryValue);
        void save([...selected]);
      });
      wrapper.append(box, document.createTextNode(view.entries[index] ?? entryValue));
      group.append(wrapper);
    }
    row.append(group);
  } else {
    const input = document.createElement('input');
    input.className = 'text-input';
    input.type = isSecretKey(view) ? 'password' : 'text';
    input.value = typeof view.value === 'string' ? view.value : '';
    // Commit on blur/Enter rather than per keystroke; each save round-trips
    // to the extension and may trigger a login.
    input.addEventListener('change', () => void save(input.value));
    row.append(input);
  }

  field.append(row, state);
  return field;
}

export function renderPreferences(
  container: HTMLElement,
  views: SourcePreferenceView[],
  commit: PreferenceCommit,
): void {
  container.replaceChildren(...views.map((view) => renderPreferenceField(container, view, commit)));
  if (views.length === 0) {
    const empty = document.createElement('p');
    empty.className = 'field-summary';
    empty.textContent = 'This source has no settings.';
    container.append(empty);
  }
}

/** Shown in place of the fields when the source picker is on "All sources". */
export function renderPreferencesUnavailable(container: HTMLElement, message: string): void {
  const note = document.createElement('p');
  note.className = 'field-summary';
  note.textContent = message;
  container.replaceChildren(note);
}
