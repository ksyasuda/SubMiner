import assert from 'node:assert/strict';
import test from 'node:test';
import {
  resolveOverlayReadinessNotificationType,
  shouldShowDesktop,
  shouldShowOverlay,
  shouldShowOsd,
} from './notification-routing';

test('notification routing preserves system notification while overlay is not ready', () => {
  assert.equal(resolveOverlayReadinessNotificationType('system', false), 'system');
});

test('notification routing preserves both as osd plus system while overlay is not ready', () => {
  assert.equal(resolveOverlayReadinessNotificationType('both', false), 'osd-system');
});

test('notification routing falls back overlay-only notification to osd while overlay is not ready', () => {
  assert.equal(resolveOverlayReadinessNotificationType('overlay', false), 'osd');
});

test('notification routing predicates classify delivery channels', () => {
  assert.equal(shouldShowOverlay('both'), true);
  assert.equal(shouldShowOverlay('system'), false);
  assert.equal(shouldShowOsd('osd-system'), true);
  assert.equal(shouldShowOsd('both'), false);
  assert.equal(shouldShowDesktop('osd-system'), true);
  assert.equal(shouldShowDesktop('overlay'), false);
});
