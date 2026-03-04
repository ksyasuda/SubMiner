import test from 'node:test';
import assert from 'node:assert/strict';
import { createResolveContext } from './context';
import { applySubtitleDomainConfig } from './subtitle-domains';

test('subtitleStyle preserveLineBreaks falls back while merge is preserved', () => {
  const { context, warnings } = createResolveContext({
    subtitleStyle: {
      preserveLineBreaks: 'invalid' as unknown as boolean,
      backgroundColor: 'rgb(1, 2, 3, 0.5)',
      secondary: {
        fontColor: 'yellow',
      },
    },
  });

  applySubtitleDomainConfig(context);

  assert.equal(context.resolved.subtitleStyle.preserveLineBreaks, false);
  assert.equal(context.resolved.subtitleStyle.backgroundColor, 'rgb(1, 2, 3, 0.5)');
  assert.equal(context.resolved.subtitleStyle.secondary.fontColor, 'yellow');
  assert.ok(
    warnings.some(
      (warning) =>
        warning.path === 'subtitleStyle.preserveLineBreaks' &&
        warning.message === 'Expected boolean.',
    ),
  );
});

test('subtitleStyle autoPauseVideoOnHover falls back on invalid value', () => {
  const { context, warnings } = createResolveContext({
    subtitleStyle: {
      autoPauseVideoOnHover: 'invalid' as unknown as boolean,
    },
  });

  applySubtitleDomainConfig(context);

  assert.equal(context.resolved.subtitleStyle.autoPauseVideoOnHover, true);
  assert.ok(
    warnings.some(
      (warning) =>
        warning.path === 'subtitleStyle.autoPauseVideoOnHover' &&
        warning.message === 'Expected boolean.',
    ),
  );
});

test('subtitleStyle autoPauseVideoOnYomitanPopup falls back on invalid value', () => {
  const { context, warnings } = createResolveContext({
    subtitleStyle: {
      autoPauseVideoOnYomitanPopup: 'invalid' as unknown as boolean,
    },
  });

  applySubtitleDomainConfig(context);

  assert.equal(context.resolved.subtitleStyle.autoPauseVideoOnYomitanPopup, false);
  assert.ok(
    warnings.some(
      (warning) =>
        warning.path === 'subtitleStyle.autoPauseVideoOnYomitanPopup' &&
        warning.message === 'Expected boolean.',
    ),
  );
});

test('subtitleStyle frequencyDictionary.matchMode accepts valid values and warns on invalid', () => {
  const valid = createResolveContext({
    subtitleStyle: {
      frequencyDictionary: {
        matchMode: 'surface',
      },
    },
  });
  applySubtitleDomainConfig(valid.context);
  assert.equal(valid.context.resolved.subtitleStyle.frequencyDictionary.matchMode, 'surface');

  const invalid = createResolveContext({
    subtitleStyle: {
      frequencyDictionary: {
        matchMode: 'reading' as unknown as 'headword' | 'surface',
      },
    },
  });
  applySubtitleDomainConfig(invalid.context);
  assert.equal(invalid.context.resolved.subtitleStyle.frequencyDictionary.matchMode, 'headword');
  assert.ok(
    invalid.warnings.some(
      (warning) =>
        warning.path === 'subtitleStyle.frequencyDictionary.matchMode' &&
        warning.message === "Expected 'headword' or 'surface'.",
    ),
  );
});
