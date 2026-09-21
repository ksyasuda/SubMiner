import assert from 'node:assert/strict';
import test from 'node:test';

import { resolveAnimatedImageLeadInSeconds, extractSoundFilenames } from './animated-image-sync';

test('extractSoundFilenames returns ordered sound filenames from an Anki field value', () => {
  assert.deepEqual(extractSoundFilenames('before [sound:word.mp3] middle [sound:alt.ogg] after'), [
    'word.mp3',
    'alt.ogg',
  ]);
});

test('resolveAnimatedImageLeadInSeconds sums configured word audio durations for animated images', async () => {
  const leadInSeconds = await resolveAnimatedImageLeadInSeconds({
    config: {
      fields: {
        audio: 'SentenceAudio',
        wordAudio: 'Pronunciation',
      },
      media: {
        imageType: 'avif',
        syncAnimatedImageToWordAudio: true,
        audioPadding: 0,
      },
    },
    noteInfo: {
      noteId: 42,
      fields: {
        SentenceAudio: { value: '[sound:sentence.mp3]' },
        Pronunciation: {
          value: '[sound:word.mp3][sound:alt.ogg]',
        },
      },
    },
    resolveConfiguredFieldName: (noteInfo, ...preferredNames) => {
      for (const preferredName of preferredNames) {
        if (!preferredName) continue;
        const resolved = Object.keys(noteInfo.fields).find(
          (fieldName) => fieldName.toLowerCase() === preferredName.toLowerCase(),
        );
        if (resolved) return resolved;
      }
      return null;
    },
    retrieveMediaFileBase64: async (filename) =>
      filename === 'word.mp3' ? 'd29yZA==' : filename === 'alt.ogg' ? 'YWx0' : '',
    probeAudioDurationSeconds: async (_buffer, filename) =>
      filename === 'word.mp3' ? 0.41 : filename === 'alt.ogg' ? 0.84 : null,
    logWarn: () => undefined,
  });

  assert.equal(leadInSeconds, 1.25);
});

test('resolveAnimatedImageLeadInSeconds does not double-count sentence audio padding', async () => {
  const leadInSeconds = await resolveAnimatedImageLeadInSeconds({
    config: {
      fields: {
        audio: 'ExpressionAudio',
      },
      media: {
        imageType: 'avif',
        syncAnimatedImageToWordAudio: true,
        audioPadding: 0.5,
      },
    },
    noteInfo: {
      noteId: 42,
      fields: {
        ExpressionAudio: {
          value: '[sound:word.mp3][sound:alt.ogg]',
        },
      },
    },
    resolveConfiguredFieldName: (noteInfo, ...preferredNames) => {
      for (const preferredName of preferredNames) {
        if (!preferredName) continue;
        const resolved = Object.keys(noteInfo.fields).find(
          (fieldName) => fieldName.toLowerCase() === preferredName.toLowerCase(),
        );
        if (resolved) return resolved;
      }
      return null;
    },
    retrieveMediaFileBase64: async (filename) =>
      filename === 'word.mp3' ? 'd29yZA==' : filename === 'alt.ogg' ? 'YWx0' : '',
    probeAudioDurationSeconds: async (_buffer, filename) =>
      filename === 'word.mp3' ? 0.41 : filename === 'alt.ogg' ? 0.84 : null,
    logWarn: () => undefined,
  });

  assert.equal(leadInSeconds, 1.25);
});

test('resolveAnimatedImageLeadInSeconds falls back to zero when sync is disabled', async () => {
  const leadInSeconds = await resolveAnimatedImageLeadInSeconds({
    config: {
      fields: {
        audio: 'ExpressionAudio',
      },
      media: {
        imageType: 'avif',
        syncAnimatedImageToWordAudio: false,
      },
    },
    noteInfo: {
      noteId: 42,
      fields: {
        ExpressionAudio: {
          value: '[sound:word.mp3]',
        },
      },
    },
    resolveConfiguredFieldName: () => 'ExpressionAudio',
    retrieveMediaFileBase64: async () => {
      throw new Error('should not be called');
    },
    probeAudioDurationSeconds: async () => {
      throw new Error('should not be called');
    },
    logWarn: () => undefined,
  });

  assert.equal(leadInSeconds, 0);
});

for (const sentenceAudio of ['', '[sound:sentence.mp3]']) {
  test(`word audio defaults independently of sentence audio (${sentenceAudio ? 'existing' : 'new'} note)`, async () => {
    const retrieved: string[] = [];
    const leadInSeconds = await resolveAnimatedImageLeadInSeconds({
      config: {
        fields: { audio: 'SentenceAudio' },
        media: { imageType: 'avif' },
      },
      noteInfo: {
        noteId: 42,
        fields: {
          ExpressionAudio: { value: '[sound:word.mp3]' },
          SentenceAudio: { value: sentenceAudio },
        },
      },
      resolveConfiguredFieldName: (noteInfo, ...preferredNames) =>
        preferredNames.find((name) => name !== undefined && name in noteInfo.fields) ?? null,
      retrieveMediaFileBase64: async (filename) => {
        retrieved.push(filename);
        return 'd29yZA==';
      },
      probeAudioDurationSeconds: async (_buffer, filename) => (filename === 'word.mp3' ? 0.6 : 4),
    });

    assert.equal(leadInSeconds, 0.6);
    assert.deepEqual(retrieved, ['word.mp3']);
  });
}
