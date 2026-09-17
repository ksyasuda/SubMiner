import type { BridgePreference } from './types';

/**
 * Extension preferences arrive as Android preference objects, one wrapper key
 * per widget type. They are stored and sent back verbatim so the extension sees
 * exactly the shape it produced; only the value field is edited.
 */

export type PreferenceKind = 'text' | 'list' | 'multi' | 'switch';

/** A preference flattened for rendering. */
export interface SourcePreferenceView {
  key: string;
  kind: PreferenceKind;
  title: string;
  summary: string | null;
  /** Current value: string for text/list, string[] for multi, boolean for switch. */
  value: string | string[] | boolean;
  /** Display labels, parallel to entryValues. Empty for text/switch. */
  entries: string[];
  entryValues: string[];
}

const WIDGETS = {
  editTextPreference: 'text',
  listPreference: 'list',
  multiSelectListPreference: 'multi',
  switchPreferenceCompat: 'switch',
  checkBoxPreference: 'switch',
} as const satisfies Record<string, PreferenceKind>;

type WidgetName = keyof typeof WIDGETS;

function widgetOf(
  entry: BridgePreference,
): { name: WidgetName; body: Record<string, unknown> } | null {
  for (const name of Object.keys(WIDGETS) as WidgetName[]) {
    const body = entry[name];
    if (body !== null && typeof body === 'object') {
      return { name, body: body as Record<string, unknown> };
    }
  }
  return null;
}

function stringList(value: unknown): string[] {
  return Array.isArray(value)
    ? value.filter((item): item is string => typeof item === 'string')
    : [];
}

/** Flatten bridge preference entries for display. Unknown widgets are skipped. */
export function parsePreferences(raw: BridgePreference[]): SourcePreferenceView[] {
  const views: SourcePreferenceView[] = [];

  for (const entry of raw) {
    if (typeof entry.key !== 'string' || entry.key.startsWith('__')) continue;
    const widget = widgetOf(entry);
    if (!widget) continue;

    const { body } = widget;
    const kind = WIDGETS[widget.name];
    const entries = stringList(body.entries);
    const entryValues = stringList(body.entryValues);

    let value: string | string[] | boolean;
    if (kind === 'multi') {
      value = stringList(body.values);
    } else if (kind === 'switch') {
      value = body.value === true;
    } else if (kind === 'list') {
      const index = typeof body.valueIndex === 'number' ? body.valueIndex : -1;
      // valueIndex is -1 when the extension has no selection yet.
      value = index >= 0 && index < entryValues.length ? entryValues[index]! : '';
    } else {
      value = typeof body.value === 'string' ? body.value : '';
    }

    views.push({
      key: entry.key,
      kind,
      title: typeof body.title === 'string' ? body.title : entry.key,
      summary: typeof body.summary === 'string' ? body.summary : null,
      value,
      entries,
      entryValues,
    });
  }

  return views;
}

/**
 * Return a copy of `raw` with one preference's value replaced, in whichever
 * fields that widget type reads. Unknown keys are returned unchanged.
 */
export function applyPreferenceValue(
  raw: BridgePreference[],
  key: string,
  value: string | string[] | boolean,
): BridgePreference[] {
  return raw.map((entry) => {
    if (entry.key !== key) return entry;
    const widget = widgetOf(entry);
    if (!widget) return entry;

    const body = { ...widget.body };
    const kind = WIDGETS[widget.name];

    if (kind === 'multi') {
      body.values = Array.isArray(value) ? value : [];
    } else if (kind === 'switch') {
      body.value = value === true;
    } else if (kind === 'list') {
      const entryValues = stringList(body.entryValues);
      const index = entryValues.indexOf(String(value));
      body.valueIndex = index;
      if (index >= 0) body.value = entryValues[index];
    } else {
      // editTextPreference carries the same string in both value and text.
      body.value = String(value);
      body.text = String(value);
    }

    return { ...entry, [widget.name]: body };
  });
}

/** True when the preference should be masked in the UI and in logs. */
export function isSecretPreference(view: SourcePreferenceView): boolean {
  return /password|token|api[-_ ]?key|secret/i.test(`${view.key} ${view.title}`);
}
