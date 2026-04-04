import test from 'node:test';
import assert from 'node:assert/strict';
import { createResolveContext } from './context';
import { applySubtitleDomainConfig } from './subtitle-domains';

test('subtitleSidebar resolves valid values and preserves dedicated defaults', () => {
  const { context } = createResolveContext({
    subtitleSidebar: {
      enabled: true,
      autoOpen: true,
      layout: 'embedded',
      toggleKey: 'KeyB',
      pauseVideoOnHover: true,
      autoScroll: false,
      maxWidth: 540,
      opacity: 0.72,
      backgroundColor: 'rgba(36, 39, 58, 0.72)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  });

  applySubtitleDomainConfig(context);

  assert.equal(context.resolved.subtitleSidebar.enabled, true);
  assert.equal(context.resolved.subtitleSidebar.autoOpen, true);
  assert.equal(context.resolved.subtitleSidebar.layout, 'embedded');
  assert.equal(context.resolved.subtitleSidebar.toggleKey, 'KeyB');
  assert.equal(context.resolved.subtitleSidebar.pauseVideoOnHover, true);
  assert.equal(context.resolved.subtitleSidebar.autoScroll, false);
  assert.equal(context.resolved.subtitleSidebar.maxWidth, 540);
  assert.equal(context.resolved.subtitleSidebar.opacity, 0.72);
  assert.equal(context.resolved.subtitleSidebar.fontFamily, '"Iosevka Aile", sans-serif');
  assert.equal(context.resolved.subtitleSidebar.fontSize, 17);
});

test('subtitleSidebar accepts zero opacity', () => {
  const { context, warnings } = createResolveContext({
    subtitleSidebar: {
      opacity: 0,
    },
  });

  applySubtitleDomainConfig(context);

  assert.equal(context.resolved.subtitleSidebar.opacity, 0);
  assert.equal(
    warnings.some((warning) => warning.path === 'subtitleSidebar.opacity'),
    false,
  );
});

test('subtitleSidebar falls back and warns on invalid values', () => {
  const { context, warnings } = createResolveContext({
    subtitleSidebar: {
      enabled: 'yes' as never,
      autoOpen: 'yes' as never,
      layout: 'floating' as never,
      maxWidth: -1,
      opacity: 5,
      fontSize: 0,
      textColor: 'blue',
      backgroundColor: 'not-a-color',
    },
  });

  applySubtitleDomainConfig(context);

  assert.equal(context.resolved.subtitleSidebar.enabled, true);
  assert.equal(context.resolved.subtitleSidebar.autoOpen, false);
  assert.equal(context.resolved.subtitleSidebar.layout, 'overlay');
  assert.equal(context.resolved.subtitleSidebar.maxWidth, 420);
  assert.equal(context.resolved.subtitleSidebar.opacity, 0.95);
  assert.equal(context.resolved.subtitleSidebar.fontSize, 16);
  assert.equal(context.resolved.subtitleSidebar.textColor, '#cad3f5');
  assert.equal(context.resolved.subtitleSidebar.backgroundColor, 'rgba(73, 77, 100, 0.9)');
  assert.ok(warnings.some((warning) => warning.path === 'subtitleSidebar.enabled'));
  assert.ok(warnings.some((warning) => warning.path === 'subtitleSidebar.autoOpen'));
  assert.ok(warnings.some((warning) => warning.path === 'subtitleSidebar.layout'));
  assert.ok(warnings.some((warning) => warning.path === 'subtitleSidebar.maxWidth'));
  assert.ok(warnings.some((warning) => warning.path === 'subtitleSidebar.opacity'));
  assert.ok(warnings.some((warning) => warning.path === 'subtitleSidebar.fontSize'));
  assert.ok(warnings.some((warning) => warning.path === 'subtitleSidebar.textColor'));
  assert.ok(
    warnings.some(
      (warning) =>
        warning.path === 'subtitleSidebar.backgroundColor' &&
        warning.message === 'Expected valid CSS color.',
    ),
  );
});
