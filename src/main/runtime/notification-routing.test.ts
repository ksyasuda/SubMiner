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

test('notification routing preserves both while overlay is not ready', () => {
  assert.equal(resolveOverlayReadinessNotificationType('both', false), 'both');
});

test('notification routing preserves overlay-only notification while overlay is not ready', () => {
  assert.equal(resolveOverlayReadinessNotificationType('overlay', false), 'overlay');
});

test('notification routing predicates classify delivery channels', () => {
  assert.equal(shouldShowOverlay('both'), true);
  assert.equal(shouldShowOverlay('system'), false);
  assert.equal(shouldShowOsd('osd-system'), true);
  assert.equal(shouldShowOsd('both'), false);
  assert.equal(shouldShowDesktop('osd-system'), true);
  assert.equal(shouldShowDesktop('overlay'), false);
});
