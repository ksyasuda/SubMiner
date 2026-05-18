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

test('subtitleStyle css declarations accept string declaration maps and warn on invalid values', () => {
  const valid = createResolveContext({
    subtitleStyle: {
      css: {
        'font-size': '42px',
        'text-wrap': 'balance',
        '--subtitle-hover-token-color': '#c6a0f6',
        '--subtitle-hover-token-background-color': 'transparent',
      },
      secondary: {
        css: {
          'text-transform': 'uppercase',
        },
      },
    },
  });
  applySubtitleDomainConfig(valid.context);
  assert.deepEqual(valid.context.resolved.subtitleStyle.css, {
    'font-size': '42px',
    'text-wrap': 'balance',
    '--subtitle-hover-token-color': '#c6a0f6',
    '--subtitle-hover-token-background-color': 'transparent',
  });
  assert.equal(valid.context.resolved.subtitleStyle.hoverTokenColor, '#c6a0f6');
  assert.equal(valid.context.resolved.subtitleStyle.hoverTokenBackgroundColor, 'transparent');
  assert.deepEqual(valid.context.resolved.subtitleStyle.secondary.css, {
    'text-transform': 'uppercase',
  });

  const invalid = createResolveContext({
    subtitleStyle: {
      css: {
        'font-size': 42,
      } as never,
      secondary: {
        css: 'font-size: 28px;' as never,
      },
    },
  });
  applySubtitleDomainConfig(invalid.context);
  assert.deepEqual(invalid.context.resolved.subtitleStyle.css, {});
  assert.deepEqual(invalid.context.resolved.subtitleStyle.secondary.css, {});
  assert.ok(invalid.warnings.some((warning) => warning.path === 'subtitleStyle.css'));
  assert.ok(invalid.warnings.some((warning) => warning.path === 'subtitleStyle.secondary.css'));
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

  assert.equal(context.resolved.subtitleStyle.autoPauseVideoOnYomitanPopup, true);
  assert.ok(
    warnings.some(
      (warning) =>
        warning.path === 'subtitleStyle.autoPauseVideoOnYomitanPopup' &&
        warning.message === 'Expected boolean.',
    ),
  );
});

test('subtitleStyle primaryDefaultMode accepts valid values and warns on invalid', () => {
  const valid = createResolveContext({
    subtitleStyle: {
      primaryDefaultMode: 'hover',
    },
  });
  applySubtitleDomainConfig(valid.context);
  assert.equal(valid.context.resolved.subtitleStyle.primaryDefaultMode, 'hover');

  const invalid = createResolveContext({
    subtitleStyle: {
      primaryDefaultMode: 'auto' as never,
    },
  });
  applySubtitleDomainConfig(invalid.context);
  assert.equal(invalid.context.resolved.subtitleStyle.primaryDefaultMode, 'visible');
  assert.ok(
    invalid.warnings.some(
      (warning) =>
        warning.path === 'subtitleStyle.primaryDefaultMode' &&
        warning.message === 'Expected hidden, visible, or hover.',
    ),
  );
});

test('subtitleStyle nameMatchEnabled falls back on invalid value', () => {
  const { context, warnings } = createResolveContext({
    subtitleStyle: {
      nameMatchEnabled: 'invalid' as unknown as boolean,
    },
  });

  applySubtitleDomainConfig(context);

  assert.equal(context.resolved.subtitleStyle.nameMatchEnabled, true);
  assert.ok(
    warnings.some(
      (warning) =>
        warning.path === 'subtitleStyle.nameMatchEnabled' &&
        warning.message === 'Expected boolean.',
    ),
  );
});

test('subtitleStyle frequencyDictionary defaults to the teal fourth band color', () => {
  const { context } = createResolveContext({});

  applySubtitleDomainConfig(context);

  assert.deepEqual(context.resolved.subtitleStyle.frequencyDictionary.bandedColors, [
    '#ed8796',
    '#f5a97f',
    '#f9e2af',
    '#8bd5ca',
    '#8aadf4',
  ]);
});

test('subtitleStyle nameMatchColor accepts valid values and warns on invalid', () => {
  const valid = createResolveContext({
    subtitleStyle: {
      nameMatchColor: '#f5bde6',
    },
  });
  applySubtitleDomainConfig(valid.context);
  assert.equal(
    (valid.context.resolved.subtitleStyle as { nameMatchColor?: string }).nameMatchColor,
    '#f5bde6',
  );

  const invalid = createResolveContext({
    subtitleStyle: {
      nameMatchColor: 'pink',
    },
  });
  applySubtitleDomainConfig(invalid.context);
  assert.equal(
    (invalid.context.resolved.subtitleStyle as { nameMatchColor?: string }).nameMatchColor,
    '#f5bde6',
  );
  assert.ok(
    invalid.warnings.some(
      (warning) =>
        warning.path === 'subtitleStyle.nameMatchColor' &&
        warning.message === 'Expected hex color.',
    ),
  );
});

test('subtitleStyle knownWordColor accepts valid values and warns on invalid', () => {
  const valid = createResolveContext({
    subtitleStyle: {
      knownWordColor: '#ed8796',
    },
  });
  applySubtitleDomainConfig(valid.context);
  assert.equal(valid.context.resolved.subtitleStyle.knownWordColor, '#ed8796');

  const invalid = createResolveContext({
    subtitleStyle: {
      knownWordColor: 'pink',
    },
  });
  applySubtitleDomainConfig(invalid.context);
  assert.equal(invalid.context.resolved.subtitleStyle.knownWordColor, '#a6da95');
  assert.ok(
    invalid.warnings.some(
      (warning) =>
        warning.path === 'subtitleStyle.knownWordColor' &&
        warning.message === 'Expected hex color.',
    ),
  );
});

test('subtitleStyle nPlusOneColor accepts valid values and warns on invalid', () => {
  const valid = createResolveContext({
    subtitleStyle: {
      nPlusOneColor: '#ed8796',
    },
  });
  applySubtitleDomainConfig(valid.context);
  assert.equal(valid.context.resolved.subtitleStyle.nPlusOneColor, '#ed8796');

  const invalid = createResolveContext({
    subtitleStyle: {
      nPlusOneColor: 'pink',
    },
  });
  applySubtitleDomainConfig(invalid.context);
  assert.equal(invalid.context.resolved.subtitleStyle.nPlusOneColor, '#c6a0f6');
  assert.ok(
    invalid.warnings.some(
      (warning) =>
        warning.path === 'subtitleStyle.nPlusOneColor' && warning.message === 'Expected hex color.',
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
