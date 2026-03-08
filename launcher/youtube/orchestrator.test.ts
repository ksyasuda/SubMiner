import test from 'node:test';
import assert from 'node:assert/strict';

import { planYoutubeSubtitleGeneration } from './orchestrator';

test('planYoutubeSubtitleGeneration prefers manual subtitles and never schedules auto-subs', () => {
  assert.deepEqual(
    planYoutubeSubtitleGeneration({
      hasPrimaryManualSubtitle: true,
      hasSecondaryManualSubtitle: false,
      secondaryCanTranslate: true,
    }),
    {
      fetchManualSubtitles: true,
      fetchAutoSubtitles: false,
      publishPrimaryManualSubtitle: false,
      publishSecondaryManualSubtitle: false,
      generatePrimarySubtitle: false,
      generateSecondarySubtitle: true,
    },
  );
});

test('planYoutubeSubtitleGeneration generates only missing tracks', () => {
  assert.deepEqual(
    planYoutubeSubtitleGeneration({
      hasPrimaryManualSubtitle: false,
      hasSecondaryManualSubtitle: true,
      secondaryCanTranslate: true,
    }),
    {
      fetchManualSubtitles: true,
      fetchAutoSubtitles: false,
      publishPrimaryManualSubtitle: false,
      publishSecondaryManualSubtitle: false,
      generatePrimarySubtitle: true,
      generateSecondarySubtitle: false,
    },
  );
});

test('planYoutubeSubtitleGeneration reuses manual tracks already present on the YouTube video', () => {
  assert.deepEqual(
    planYoutubeSubtitleGeneration({
      hasPrimaryManualSubtitle: true,
      hasSecondaryManualSubtitle: true,
      secondaryCanTranslate: true,
    }),
    {
      fetchManualSubtitles: true,
      fetchAutoSubtitles: false,
      publishPrimaryManualSubtitle: false,
      publishSecondaryManualSubtitle: false,
      generatePrimarySubtitle: false,
      generateSecondarySubtitle: false,
    },
  );
});
