import assert from 'node:assert/strict';
import test from 'node:test';

import { renderControl } from './settings-controls';
import type { ConfigSettingsField } from '../types/settings';

class FakeOption {
  value = '';
  textContent: string | null = null;
  selected = false;
  className = '';
}

class FakeSelect {
  value = '';
  className = '';
  options: FakeOption[] = [];
  private readonly listeners = new Map<string, Array<() => void>>();

  append(...children: FakeOption[]): void {
    for (const child of children) {
      this.options.push(child);
      if (child.selected) {
        this.value = child.value;
      }
    }
  }

  addEventListener(type: string, listener: () => void): void {
    const listeners = this.listeners.get(type) ?? [];
    listeners.push(listener);
    this.listeners.set(type, listeners);
  }

  dispatchEvent(event: Event): boolean {
    for (const listener of this.listeners.get(event.type) ?? []) {
      listener();
    }
    return true;
  }
}

function installDocumentStub(): () => void {
  const previousDocument = globalThis.document;
  globalThis.document = {
    createElement(tagName: string) {
      return tagName === 'select' ? new FakeSelect() : new FakeOption();
    },
  } as unknown as Document;
  return () => {
    globalThis.document = previousDocument;
  };
}

function createSelectField(): ConfigSettingsField {
  return {
    id: 'updates.notificationType',
    label: 'Notification Type',
    description: 'How SubMiner announces available updates.',
    configPath: 'updates.notificationType',
    category: 'tracking-app',
    section: 'Updates',
    control: 'select',
    defaultValue: 'system',
    enumValues: ['overlay', 'system', 'both', 'none'],
    restartBehavior: 'restart',
  };
}

test('select controls show config-only current values without offering them otherwise', () => {
  const restoreDocument = installDocumentStub();
  const updates: Array<{ path: string; value: unknown }> = [];
  try {
    const control = renderControl(createSelectField(), {
      valueForField: () => 'osd-system',
      valueForPath: () => undefined,
      updateDraft: (path, value) => updates.push({ path, value }),
      resetDraftPath: () => {},
      setFieldError: () => {},
    }) as HTMLSelectElement;

    assert.equal(control.value, 'osd-system');
    assert.deepEqual(
      Array.from(control.options).map((option) => option.value),
      ['osd-system', 'overlay', 'system', 'both', 'none'],
    );

    control.value = 'overlay';
    control.dispatchEvent(new Event('change'));

    assert.deepEqual(updates, [{ path: 'updates.notificationType', value: 'overlay' }]);
  } finally {
    restoreDocument();
  }
});
