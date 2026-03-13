import assert from 'node:assert/strict';
import test from 'node:test';

import { createControllerConfigForm } from './controller-config-form.js';

function createClassList(initialTokens: string[] = []) {
  const tokens = new Set(initialTokens);
  return {
    add: (...entries: string[]) => {
      for (const entry of entries) tokens.add(entry);
    },
    remove: (...entries: string[]) => {
      for (const entry of entries) tokens.delete(entry);
    },
    toggle: (entry: string, force?: boolean) => {
      if (force === undefined) {
        if (tokens.has(entry)) tokens.delete(entry);
        else tokens.add(entry);
        return tokens.has(entry);
      }
      if (force) tokens.add(entry);
      else tokens.delete(entry);
      return force;
    },
    contains: (entry: string) => tokens.has(entry),
  };
}

function createFakeElement() {
  return {
    className: '',
    textContent: '',
    innerHTML: '',
    children: [] as any[],
    listeners: new Map<string, Array<() => void>>(),
    classList: createClassList(),
    appendChild(child: any) {
      this.children.push(child);
      return child;
    },
    addEventListener(type: string, listener: () => void) {
      const existing = this.listeners.get(type) ?? [];
      existing.push(listener);
      this.listeners.set(type, existing);
    },
    dispatch(type: string) {
      for (const listener of this.listeners.get(type) ?? []) {
        listener();
      }
    },
  };
}

test('controller config form renders rows and dispatches learn clear reset callbacks', () => {
  const globals = globalThis as typeof globalThis & { document?: unknown };
  const previousDocument = globals.document;
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createFakeElement(),
    },
  });

  try {
    const calls: string[] = [];
    const container = createFakeElement();
    const form = createControllerConfigForm({
      container: container as never,
      getBindings: () =>
        ({
          toggleLookup: { kind: 'button', buttonIndex: 0 },
          closeLookup: { kind: 'button', buttonIndex: 1 },
          toggleKeyboardOnlyMode: { kind: 'button', buttonIndex: 3 },
          mineCard: { kind: 'button', buttonIndex: 2 },
          quitMpv: { kind: 'button', buttonIndex: 6 },
          previousAudio: { kind: 'none' },
          nextAudio: { kind: 'button', buttonIndex: 5 },
          playCurrentAudio: { kind: 'button', buttonIndex: 4 },
          toggleMpvPause: { kind: 'button', buttonIndex: 9 },
          leftStickHorizontal: { kind: 'axis', axisIndex: 0, dpadFallback: 'horizontal' },
          leftStickVertical: { kind: 'axis', axisIndex: 1, dpadFallback: 'vertical' },
          rightStickHorizontal: { kind: 'axis', axisIndex: 3, dpadFallback: 'none' },
          rightStickVertical: { kind: 'axis', axisIndex: 4, dpadFallback: 'none' },
        }) as never,
      getLearningActionId: () => 'toggleLookup',
      onLearn: (actionId) => calls.push(`learn:${actionId}`),
      onClear: (actionId) => calls.push(`clear:${actionId}`),
      onReset: (actionId) => calls.push(`reset:${actionId}`),
    });

    form.render();

    const firstRow = container.children[1];
    assert.equal(firstRow.classList.contains('learning'), true);
    assert.match(firstRow.children[1].textContent, /Button 0/);

    firstRow.children[2].children[0].dispatch('click');
    firstRow.children[2].children[1].dispatch('click');
    firstRow.children[2].children[2].dispatch('click');

    assert.deepEqual(calls, [
      'learn:toggleLookup',
      'clear:toggleLookup',
      'reset:toggleLookup',
    ]);
  } finally {
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});
