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
