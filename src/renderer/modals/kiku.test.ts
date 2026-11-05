import assert from 'node:assert/strict';
import test from 'node:test';

import type { KikuDuplicateCardInfo, KikuFieldGroupingChoice } from '../../types';
import { createKikuModal, formatKikuScrollHint, KIKU_CONTROLLER_SCROLL_STEP_PX } from './kiku';

class FakeClassList {
  private readonly names = new Set<string>();
  add(name: string): void {
    this.names.add(name);
  }
  remove(name: string): void {
    this.names.delete(name);
  }
  toggle(name: string, force?: boolean): void {
    if (force ?? !this.names.has(name)) this.names.add(name);
    else this.names.delete(name);
  }
  contains(name: string): boolean {
    return this.names.has(name);
  }
}

class FakeElement {
  classList = new FakeClassList();
  dataset: Record<string, string | undefined> = {};
  textContent = '';
  checked = false;
  disabled = false;
  scrollTop = 0;
  setAttribute(): void {}
  addEventListener(): void {}
}

const CARD = (noteId: number): KikuDuplicateCardInfo => ({
  noteId,
  expression: '猫',
  sentencePreview: '',
  hasAudio: true,
  hasImage: false,
  isOriginal: noteId === 1,
});

function setup() {
  const dom = new Proxy({} as Record<string, FakeElement>, {
    get: (target, key: string) => (target[key] ??= new FakeElement()),
  });
  const state = {
    kikuModalOpen: false,
    kikuSelectedCard: 1 as 1 | 2,
    kikuOriginalData: null,
    kikuDuplicateData: null,
    kikuPendingChoice: null,
    kikuPreviewCompactData: null,
    kikuPreviewFullData: null,
    kikuPreviewMode: 'compact' as 'compact' | 'full',
    kikuModalStep: 'select' as 'select' | 'preview',
    isOverSubtitle: false,
  };
  const responses: KikuFieldGroupingChoice[] = [];
  const previousWindow = Object.getOwnPropertyDescriptor(globalThis, 'window');
  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        kikuBuildMergePreview: async () => ({ ok: true, compact: { a: 1 }, full: { a: 1, b: 2 } }),
        kikuFieldGroupingRespond: (choice: KikuFieldGroupingChoice) => responses.push(choice),
        notifyOverlayModalClosed: () => {},
      },
    },
  });
  const modal = createKikuModal({ state, dom } as never, {
    modalStateReader: { isAnyModalOpen: () => false },
    syncSettingsModalSubtitleSuppression: () => {},
  });
  modal.openKikuFieldGroupingModal({ original: CARD(1), duplicate: CARD(2) });
  return {
    dom,
    state,
    modal,
    responses,
    restore: () => {
      if (previousWindow) Object.defineProperty(globalThis, 'window', previousWindow);
      else Reflect.deleteProperty(globalThis, 'window');
    },
  };
}

test('controller actions walk the duplicate card dialog like the keyboard does', async () => {
  const { dom, state, modal, responses, restore } = setup();
  try {
    assert.equal(modal.getControllerStep(), 'select');
    assert.equal(modal.handleControllerAction('selectNew'), true);
    assert.equal(state.kikuSelectedCard, 2);
    assert.equal(dom.kikuCard2!.classList.contains('active'), true);
    assert.equal(dom.kikuDeleteDuplicateCheckbox!.checked, true);
    modal.handleControllerAction('toggleDeleteDuplicate');
    assert.equal(dom.kikuDeleteDuplicateCheckbox!.checked, false);
    // Preview-only actions do nothing on the first step.
    assert.equal(modal.handleControllerAction('confirmMerge'), false);

    modal.handleControllerAction('continue');
    await new Promise((resolve) => setTimeout(resolve, 0));
    assert.equal(modal.getControllerStep(), 'preview');

    modal.handleControllerAction('scrollDown');
    modal.handleControllerAction('scrollDown');
    modal.handleControllerAction('scrollUp');
    assert.equal(dom.kikuPreviewJson!.scrollTop, KIKU_CONTROLLER_SCROLL_STEP_PX);
    modal.handleControllerAction('toggleView');
    assert.equal(state.kikuPreviewMode, 'full');
    assert.equal(modal.handleControllerAction('cancel'), false, 'first-step Cancel is hidden');

    modal.handleControllerAction('back');
    assert.equal(modal.getControllerStep(), 'select');
    modal.handleControllerAction('continue');
    await new Promise((resolve) => setTimeout(resolve, 0));
    modal.handleControllerAction('confirmMerge');
    assert.deepEqual(responses, [
      { keepNoteId: 2, deleteNoteId: 1, deleteDuplicate: false, cancelled: false },
    ]);
    assert.equal(state.kikuModalOpen, false);
    assert.equal(modal.handleControllerAction('continue'), false, 'closed dialog ignores input');
  } finally {
    restore();
  }
});

test('Continue is ignored while the merge preview is being built', () => {
  const { dom, modal, restore } = setup();
  try {
    dom.kikuConfirmButton!.disabled = true;
    assert.equal(modal.handleControllerAction('continue'), false);
  } finally {
    restore();
  }
});

test('controller hint tags land on the cards, buttons and preview header', () => {
  const { dom, modal, restore } = setup();
  try {
    modal.setControllerHints({
      selectOriginal: '◀',
      selectNew: '▶',
      toggleDeleteDuplicate: 'X',
      continue: 'A',
      cancel: 'B',
      confirmMerge: 'A',
      back: 'B',
      cancelMerge: null,
      toggleView: 'Y',
      scrollUp: '▲',
      scrollDown: '▼',
    });
    assert.equal(dom.kikuCard1Label!.dataset.controllerHint, '◀');
    assert.equal(dom.kikuCard2Label!.dataset.controllerHint, '▶');
    assert.equal(dom.kikuDeleteToggle!.dataset.controllerHint, 'X');
    assert.equal(dom.kikuConfirmButton!.dataset.controllerHint, 'A');
    assert.equal(dom.kikuBackButton!.dataset.controllerHint, 'B');
    assert.equal(dom.kikuFinalCancelButton!.dataset.controllerHint, undefined);
    assert.equal(dom.kikuPreviewToggle!.dataset.controllerHint, 'Y');
    assert.equal(dom.kikuPreviewScrollHint!.dataset.controllerHint, '▲ ▼');

    modal.setControllerHintsVisible(true);
    assert.equal(dom.kikuModal!.classList.contains('controller-hints-visible'), true);
  } finally {
    restore();
  }
});

test('the scroll tag shows whichever directions are bound', () => {
  assert.equal(formatKikuScrollHint('▲', '▼'), '▲ ▼');
  assert.equal(formatKikuScrollHint(null, 'RS ▼'), 'RS ▼');
  assert.equal(formatKikuScrollHint(null, null), null);
});
