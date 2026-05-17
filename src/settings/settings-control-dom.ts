import type { ConfigSettingsSnapshotValue } from '../types/settings';

export function createElement<K extends keyof HTMLElementTagNameMap>(
  tagName: K,
  className?: string,
): HTMLElementTagNameMap[K] {
  const element = document.createElement(tagName);
  if (className) {
    element.className = className;
  }
  return element;
}

export function addOption(select: HTMLSelectElement, value: string, label = value): void {
  const option = createElement('option') as HTMLOptionElement;
  option.value = value;
  option.textContent = label;
  select.append(option);
}

export function uniqueSorted(values: Iterable<string>): string[] {
  return [...new Set([...values].filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

export function isSecretSnapshotValue(
  value: ConfigSettingsSnapshotValue,
): value is { configured: boolean } {
  return Boolean(
    value &&
      typeof value === 'object' &&
      'configured' in value &&
      typeof (value as { configured?: unknown }).configured === 'boolean',
  );
}
