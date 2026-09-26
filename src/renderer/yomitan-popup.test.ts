import test from 'node:test';
import assert from 'node:assert/strict';
import {
  YOMITAN_LOOKUP_EVENT,
  YOMITAN_POPUP_VISIBLE_HOST_SELECTOR,
  isYomitanPopupVisible,
  registerYomitanLookupListener,
  registerDictionaryPopupVisibilityListener,
} from './yomitan-popup.js';

test('native popup attention events from either backend have the same lifecycle', () => {
  for (const backend of ['yomitan', 'hachidori']) {
    const target = new EventTarget();
    const calls: string[] = [];
    const disposeShown = registerDictionaryPopupVisibilityListener(
      'shown',
      () => calls.push('shown'),
      target,
    );
    const disposeHidden = registerDictionaryPopupVisibilityListener(
      'hidden',
      () => calls.push('hidden'),
      target,
    );
    target.dispatchEvent(new CustomEvent(`${backend}-popup-shown`));
    target.dispatchEvent(new CustomEvent(`${backend}-popup-hidden`));
    disposeShown();
    disposeHidden();
    target.dispatchEvent(new CustomEvent(`${backend}-popup-shown`));
    target.dispatchEvent(new CustomEvent(`${backend}-popup-hidden`));
    assert.deepEqual(calls, ['shown', 'hidden']);
  }
});

test('registerYomitanLookupListener forwards the SubMiner Yomitan lookup event', () => {
  const target = new EventTarget();
  const calls: string[] = [];

  const dispose = registerYomitanLookupListener(target, () => {
    calls.push('lookup');
  });

  target.dispatchEvent(new CustomEvent(YOMITAN_LOOKUP_EVENT));
  dispose();
  target.dispatchEvent(new CustomEvent(YOMITAN_LOOKUP_EVENT));

  assert.deepEqual(calls, ['lookup']);
});

test('isYomitanPopupVisible falls back to querySelector when querySelectorAll is unavailable', () => {
  const root = {
    querySelector: (selector: string) =>
      selector === YOMITAN_POPUP_VISIBLE_HOST_SELECTOR ? ({} as Element) : null,
  } as ParentNode;

  assert.equal(isYomitanPopupVisible(root), true);
});
