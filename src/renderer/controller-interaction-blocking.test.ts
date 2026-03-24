import assert from 'node:assert/strict';
import test from 'node:test';

import { isControllerInteractionBlocked } from './controller-interaction-blocking.js';

test('subtitle sidebar stays controller-passive while other modals block controller input', () => {
  assert.equal(
    isControllerInteractionBlocked({
      controllerSelectModalOpen: false,
      controllerDebugModalOpen: false,
      jimakuModalOpen: false,
      kikuModalOpen: false,
      runtimeOptionsModalOpen: false,
      subsyncModalOpen: false,
      youtubePickerModalOpen: false,
      sessionHelpModalOpen: false,
      subtitleSidebarModalOpen: true,
    }),
    false,
  );

  assert.equal(
    isControllerInteractionBlocked({
      controllerSelectModalOpen: false,
      controllerDebugModalOpen: false,
      jimakuModalOpen: false,
      kikuModalOpen: false,
      runtimeOptionsModalOpen: true,
      subsyncModalOpen: false,
      youtubePickerModalOpen: false,
      sessionHelpModalOpen: false,
      subtitleSidebarModalOpen: false,
    }),
    true,
  );
});
