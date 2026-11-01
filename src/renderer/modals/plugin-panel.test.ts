import assert from 'node:assert/strict';
import test from 'node:test';

import type { RendererContext } from '../context';
import { createRendererState } from '../state';
import { createPluginPanelModal } from './plugin-panel';

function createHarness() {
  let activeElement: unknown = null;
  class TestElement {
    textContent = '';
    scrollTop = 0;
    isConnected = true;
    classList = { add() {}, remove() {} };
    setAttribute() {}
    replaceChildren() {}
    addEventListener() {}
    removeEventListener() {}
    contains() {
      return true;
    }
    focus() {
      activeElement = this;
    }
    getClientRects() {
      return [{}];
    }
  }
  const originals = new Map(
    ['window', 'document', 'Element', 'HTMLElement'].map((name) => [
      name,
      Object.getOwnPropertyDescriptor(globalThis, name),
    ]),
  );
  const closed: number[] = [];
  const dom = {
    overlay: new TestElement(),
    pluginPanelModal: new TestElement(),
    pluginPanelTitle: new TestElement(),
    pluginPanelBody: new TestElement(),
    pluginPanelClose: new TestElement(),
  };
  for (const [name, value] of Object.entries({
    Element: TestElement,
    HTMLElement: TestElement,
    document: {
      get activeElement() {
        return activeElement;
      },
      addEventListener() {},
      removeEventListener() {},
    },
    window: {
      focus() {},
      addEventListener() {},
      removeEventListener() {},
      electronAPI: {
        notifyOverlayModalClosed(_modal: string, id: number) {
          closed.push(id);
        },
      },
    },
  })) {
    Object.defineProperty(globalThis, name, { configurable: true, writable: true, value });
  }
  const state = createRendererState();
  activeElement = dom.pluginPanelBody;
  const ctx = {
    dom,
    state,
    platform: { isModalLayer: true, shouldToggleMouseIgnore: false },
  } as unknown as RendererContext;
  const modal = createPluginPanelModal(ctx, {
    modalStateReader: { isAnyModalOpen: () => state.pluginPanelModalOpen },
    syncSettingsModalSubtitleSuppression: () => {},
  });
  return {
    modal,
    dom,
    state,
    closed,
    restore() {
      for (const [name, descriptor] of originals) {
        if (descriptor) Object.defineProperty(globalThis, name, descriptor);
        else Reflect.deleteProperty(globalThis, name);
      }
    },
  };
}

test('late opens cannot replace a newer panel or acknowledge a stale instance', () => {
  const h = createHarness();
  try {
    assert.equal(h.modal.openPluginPanelModal({ panelId: 2, title: 'new', markdown: '' }), true);
    assert.equal(h.modal.openPluginPanelModal({ panelId: 1, title: 'old', markdown: '' }), false);
    assert.equal(h.dom.pluginPanelTitle.textContent, 'new');
    h.modal.closePluginPanelModal();
    assert.deepEqual(h.closed, [2]);
    assert.equal(h.modal.openPluginPanelModal({ panelId: 2, title: 'retry', markdown: '' }), false);
  } finally {
    h.restore();
  }
});

test('a close received before its open prevents the late panel from appearing', () => {
  const h = createHarness();
  try {
    h.modal.closePluginPanelFromMain(1);
    assert.equal(
      h.modal.openPluginPanelModal({ panelId: 1, title: 'closed', markdown: '' }),
      false,
    );
    assert.equal(h.state.pluginPanelModalOpen, false);
    assert.equal(h.modal.openPluginPanelModal({ panelId: 2, title: 'next', markdown: '' }), true);
    h.modal.closePluginPanelFromMain(1);
    assert.equal(h.state.pluginPanelModalOpen, true);
    h.modal.closePluginPanelModal();
  } finally {
    h.restore();
  }
});
