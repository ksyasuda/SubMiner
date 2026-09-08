import assert from 'node:assert/strict';
import test from 'node:test';
import { resolveConfig } from '../resolve';
import { buildConfigSettingsRegistry } from '../settings/registry';
import { SUBTITLE_GENERATION_MODELS } from '../../shared/subtitle-generation-model-catalog';
import { resolveSubtitleGenerationConfig } from '../../shared/subtitle-generation';

test('every downloadable multilingual model is accepted by config and offered in settings', () => {
  const settings = buildConfigSettingsRegistry(resolveConfig({}).resolved).find(
    (entry) => entry.configPath === 'subtitleGeneration.managedModel',
  );
  assert.deepEqual(
    settings?.enumValues,
    SUBTITLE_GENERATION_MODELS.map((model) => model.id),
  );
  for (const { id } of SUBTITLE_GENERATION_MODELS) {
    const { resolved, warnings } = resolveConfig({ subtitleGeneration: { managedModel: id } });
    assert.equal(resolved.subtitleGeneration.managedModel, id);
    assert.equal(warnings.length, 0);
  }
});

test('generation config rejects English-only, unknown, and prototype model names', () => {
  for (const managedModel of ['tiny.en', 'small.en-q5_1', 'unknown', 'toString', '__proto__']) {
    const warnings: string[] = [];
    const resolved = resolveSubtitleGenerationConfig({ managedModel }, (key) => warnings.push(key));
    assert.equal(resolved.managedModel, 'small');
    assert.equal(warnings.length, 1);
  }
});

test('generation config is resolved and its external model path is editable without restart', () => {
  const { resolved, warnings } = resolveConfig({
    subtitleGeneration: { modelPath: '/models/japanese.bin', managedModel: 'medium', threads: 8 },
  });
  assert.equal(resolved.subtitleGeneration.modelPath, '/models/japanese.bin');
  assert.equal(resolved.subtitleGeneration.managedModel, 'medium');
  assert.equal(warnings.length, 0);
  const field = buildConfigSettingsRegistry(resolved).find(
    (entry) => entry.configPath === 'subtitleGeneration.modelPath',
  );
  assert.equal(field?.category, 'integrations');
  assert.equal(field?.restartBehavior, 'hot-reload');
});

test('generation executable overrides default to empty and accept blank values without warnings', () => {
  for (const subtitleGeneration of [
    {},
    { whisperPath: '', ffmpegPath: '', ffprobePath: '' },
    { whisperPath: '  ', ffmpegPath: '  ', ffprobePath: '  ' },
  ]) {
    const { resolved, warnings } = resolveConfig({ subtitleGeneration });
    assert.equal(resolved.subtitleGeneration.whisperPath, '');
    assert.equal(resolved.subtitleGeneration.ffmpegPath, '');
    assert.equal(resolved.subtitleGeneration.ffprobePath, '');
    assert.equal(warnings.length, 0);
  }
});

test('subtitle generation shortcut can be customized or disabled', () => {
  assert.equal(resolveConfig({}).resolved.shortcuts.openSubtitleGeneration, 'Ctrl+Shift+G');
  assert.equal(
    resolveConfig({ shortcuts: { openSubtitleGeneration: 'Ctrl+Alt+G' } }).resolved.shortcuts
      .openSubtitleGeneration,
    'Ctrl+Alt+G',
  );
  assert.equal(
    resolveConfig({ shortcuts: { openSubtitleGeneration: null } }).resolved.shortcuts
      .openSubtitleGeneration,
    null,
  );
});

test('dialogue detection paths are optional, validated, and editable in Settings', () => {
  const { resolved, warnings } = resolveConfig({
    subtitleGeneration: { vadModelPath: ' /models/silero.bin ', vadPath: ' /bin/vad ' },
  });
  assert.equal(resolved.subtitleGeneration.vadModelPath, '/models/silero.bin');
  assert.equal(resolved.subtitleGeneration.vadPath, '/bin/vad');
  assert.equal(warnings.length, 0);
  for (const key of ['vadModelPath', 'vadPath'] as const) {
    const field = buildConfigSettingsRegistry(resolved).find(
      (entry) => entry.configPath === `subtitleGeneration.${key}`,
    );
    assert.equal(field?.category, 'integrations');
    assert.equal(field?.restartBehavior, 'hot-reload');
    assert.equal(resolveConfig({}).resolved.subtitleGeneration[key], '');
    assert.equal(resolveConfig({ subtitleGeneration: { [key]: false } }).warnings.length, 1);
  }
});
