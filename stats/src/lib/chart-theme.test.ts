import assert from 'node:assert/strict';
import test from 'node:test';
import { CHART_DEFAULTS, CHART_THEME, TOOLTIP_CONTENT_STYLE } from './chart-theme';

test('CHART_THEME exposes a grid color', () => {
  assert.equal(CHART_THEME.grid, '#494d64');
});

test('CHART_DEFAULTS uses 11px ticks for legibility', () => {
  assert.equal(CHART_DEFAULTS.tickFontSize, 11);
});

test('TOOLTIP_CONTENT_STYLE mirrors the shared tooltip colors', () => {
  assert.equal(TOOLTIP_CONTENT_STYLE.background, CHART_THEME.tooltipBg);
  assert.ok(String(TOOLTIP_CONTENT_STYLE.border).includes(CHART_THEME.tooltipBorder));
});
