import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';

const rendererSource = fs.readFileSync(
  path.join(process.cwd(), 'src/renderer/renderer.ts'),
  'utf8',
);

function indexOfRequired(pattern: string): number {
  const index = rendererSource.indexOf(pattern);
  assert.notEqual(index, -1, `Expected renderer.ts to contain ${pattern}`);
  return index;
}

test('renderer applies subtitle style and position before first subtitle paint', () => {
  const styleIndex = indexOfRequired(
    'const initialSubtitleStyle = await window.electronAPI.getSubtitleStyle();',
  );
  const positionIndex = indexOfRequired(
    "await window.electronAPI.getSubtitlePosition(),\n    'startup',",
  );
  const listenerIndex = indexOfRequired('window.electronAPI.onSubtitle((data: SubtitleData) => {');
  const currentSubtitleIndex = indexOfRequired(
    'initialSubtitle = await window.electronAPI.getCurrentSubtitle();',
  );

  assert.ok(styleIndex < listenerIndex);
  assert.ok(positionIndex < listenerIndex);
  assert.ok(styleIndex < currentSubtitleIndex);
  assert.ok(positionIndex < currentSubtitleIndex);
});

test('renderer renders initial subtitle snapshot before subscribing to live subtitle updates', () => {
  const listenerIndex = indexOfRequired('window.electronAPI.onSubtitle((data: SubtitleData) => {');
  const currentSubtitleIndex = indexOfRequired(
    'initialSubtitle = await window.electronAPI.getCurrentSubtitle();',
  );
  const initialRenderIndex = indexOfRequired('subtitleRenderer.renderSubtitle(initialSubtitle);');

  assert.ok(currentSubtitleIndex < initialRenderIndex);
  assert.ok(initialRenderIndex < listenerIndex);
});

test('renderer reports subtitle bounds immediately after initial subtitle layout', () => {
  const initialRenderIndex = indexOfRequired('subtitleRenderer.renderSubtitle(initialSubtitle);');
  const initialLayoutIndex = indexOfRequired(
    'subtitleRenderer.renderSubtitle(initialSubtitle);\n  positioning.applyYPercent(positioning.getCurrentYPercent());',
  );
  const immediateMeasurementIndex = indexOfRequired(
    'positioning.applyYPercent(positioning.getCurrentYPercent());\n  measurementReporter.emitNow();',
  );
  const listenerIndex = indexOfRequired('window.electronAPI.onSubtitle((data: SubtitleData) => {');

  assert.equal(initialRenderIndex, initialLayoutIndex);
  assert.ok(initialLayoutIndex < immediateMeasurementIndex);
  assert.ok(immediateMeasurementIndex < listenerIndex);
});

test('renderer wires subtitle pointer handlers before first subtitle paint', () => {
  const primaryMouseEnterIndex = indexOfRequired(
    "ctx.dom.subtitleContainer.addEventListener('mouseenter', mouseHandlers.handlePrimaryMouseEnter);",
  );
  const pointerTrackingIndex = indexOfRequired('mouseHandlers.setupPointerTracking();');
  const initialRenderIndex = indexOfRequired('subtitleRenderer.renderSubtitle(initialSubtitle);');
  const initialMeasurementIndex = indexOfRequired(
    'positioning.applyYPercent(positioning.getCurrentYPercent());\n  measurementReporter.emitNow();',
  );

  assert.ok(primaryMouseEnterIndex < initialRenderIndex);
  assert.ok(pointerTrackingIndex < initialRenderIndex);
  assert.ok(primaryMouseEnterIndex < initialMeasurementIndex);
  assert.ok(pointerTrackingIndex < initialMeasurementIndex);
});

test('renderer reports subtitle bounds immediately after live subtitle layout', () => {
  const liveRenderIndex = indexOfRequired('subtitleRenderer.renderSubtitle(data);');
  const liveLayoutIndex = indexOfRequired(
    'subtitleRenderer.renderSubtitle(data);\n      positioning.applyYPercent(positioning.getCurrentYPercent());',
  );
  const immediateMeasurementIndex = indexOfRequired(
    'positioning.applyYPercent(positioning.getCurrentYPercent());\n      measurementReporter.emitNow();',
  );
  const sidebarUpdateIndex = indexOfRequired('subtitleSidebarModal.handleSubtitleUpdated(data);');

  assert.equal(liveRenderIndex, liveLayoutIndex);
  assert.ok(liveLayoutIndex < immediateMeasurementIndex);
  assert.ok(immediateMeasurementIndex < sidebarUpdateIndex);
});

test('renderer restores subtitle sidebar open state only on visible overlay layer', () => {
  const sidebarRestoreIndex = indexOfRequired(
    "ctx.platform.overlayLayer === 'visible' && (await window.electronAPI.getSubtitleSidebarOpen())",
  );
  const sidebarModalIndex = indexOfRequired(
    'const subtitleSidebarModal = createSubtitleSidebarModal',
  );

  assert.ok(sidebarModalIndex < sidebarRestoreIndex);
});
