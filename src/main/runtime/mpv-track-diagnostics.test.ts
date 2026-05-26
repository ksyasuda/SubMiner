import assert from 'node:assert/strict';
import test from 'node:test';
import { buildSubtitleTrackDiagnostics } from './mpv-track-diagnostics';

test('buildSubtitleTrackDiagnostics summarizes subtitle tracks without dumping track list', () => {
  const diagnostics = buildSubtitleTrackDiagnostics(3, [
    { type: 'video', id: 1, selected: true },
    { type: 'sub', id: 3, lang: 'ja', selected: true, external: false, codec: 'ass' },
    { type: 'sub', id: '4', title: 'English', external: true, codec: 'srt' },
    { type: 'audio', id: 2, lang: 'jpn' },
  ]);

  assert.deepEqual(diagnostics, {
    trackListReadable: true,
    trackCount: 4,
    subtitleTrackCount: 2,
    activePrimarySid: 3,
    selectedSubtitleIds: [3],
    externalSubtitleCount: 1,
    internalSubtitleCount: 1,
    languages: ['ja'],
    selectedSubtitleLabels: ['internal#3:ja'],
  });
});

test('buildSubtitleTrackDiagnostics marks unreadable track list', () => {
  assert.deepEqual(buildSubtitleTrackDiagnostics(null, null), {
    trackListReadable: false,
    trackCount: 0,
    subtitleTrackCount: 0,
    activePrimarySid: null,
    selectedSubtitleIds: [],
    externalSubtitleCount: 0,
    internalSubtitleCount: 0,
    languages: [],
    selectedSubtitleLabels: [],
  });
});
