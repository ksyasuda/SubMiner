import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createStartupOsdSequencer,
  type StartupOsdSequencerCharacterDictionaryEvent,
} from './startup-osd-sequencer';

function makeDictionaryEvent(
  phase: StartupOsdSequencerCharacterDictionaryEvent['phase'],
  message: string,
): StartupOsdSequencerCharacterDictionaryEvent {
  return {
    phase,
    message,
  };
}

test('startup OSD keeps dictionary progress hidden until tokenization and annotation loading finish', () => {
  const osdMessages: string[] = [];
  const sequencer = createStartupOsdSequencer({
    showOsd: (message) => {
      osdMessages.push(message);
    },
  });

  sequencer.notifyCharacterDictionaryStatus(
    makeDictionaryEvent('syncing', 'Updating character dictionary for Frieren...'),
  );
  sequencer.showAnnotationLoading('Loading subtitle annotations |');
  sequencer.markTokenizationReady();

  assert.deepEqual(osdMessages, ['Loading subtitle annotations |']);

  sequencer.showAnnotationLoading('Loading subtitle annotations /');
  assert.deepEqual(osdMessages, [
    'Loading subtitle annotations |',
    'Loading subtitle annotations /',
  ]);

  sequencer.markAnnotationLoadingComplete('Subtitle annotations loaded');
  assert.deepEqual(osdMessages, [
    'Loading subtitle annotations |',
    'Loading subtitle annotations /',
    'Updating character dictionary for Frieren...',
  ]);
});

test('startup OSD buffers checking behind annotations and replaces it with later generating progress', () => {
  const osdMessages: string[] = [];
  const sequencer = createStartupOsdSequencer({
    showOsd: (message) => {
      osdMessages.push(message);
    },
  });

  sequencer.notifyCharacterDictionaryStatus(
    makeDictionaryEvent('checking', 'Checking character dictionary for Frieren...'),
  );
  sequencer.showAnnotationLoading('Loading subtitle annotations |');
  sequencer.markTokenizationReady();
  sequencer.notifyCharacterDictionaryStatus(
    makeDictionaryEvent('generating', 'Generating character dictionary for Frieren...'),
  );

  assert.deepEqual(osdMessages, ['Loading subtitle annotations |']);

  sequencer.markAnnotationLoadingComplete('Subtitle annotations loaded');

  assert.deepEqual(osdMessages, [
    'Loading subtitle annotations |',
    'Generating character dictionary for Frieren...',
  ]);
});

test('startup OSD replaces earlier dictionary progress with later building progress', () => {
  const osdMessages: string[] = [];
  const sequencer = createStartupOsdSequencer({
    showOsd: (message) => {
      osdMessages.push(message);
    },
  });

  sequencer.notifyCharacterDictionaryStatus(
    makeDictionaryEvent('syncing', 'Updating character dictionary for Frieren...'),
  );
  sequencer.showAnnotationLoading('Loading subtitle annotations |');
  sequencer.markTokenizationReady();
  sequencer.notifyCharacterDictionaryStatus(
    makeDictionaryEvent('building', 'Building character dictionary for Frieren...'),
  );

  sequencer.markAnnotationLoadingComplete('Subtitle annotations loaded');

  assert.deepEqual(osdMessages, [
    'Loading subtitle annotations |',
    'Building character dictionary for Frieren...',
  ]);
});

test('startup OSD skips buffered dictionary ready messages when progress completed before it became visible', () => {
  const osdMessages: string[] = [];
  const sequencer = createStartupOsdSequencer({
    showOsd: (message) => {
      osdMessages.push(message);
    },
  });

  sequencer.notifyCharacterDictionaryStatus(
    makeDictionaryEvent('syncing', 'Updating character dictionary for Frieren...'),
  );
  sequencer.notifyCharacterDictionaryStatus(
    makeDictionaryEvent('ready', 'Character dictionary ready for Frieren'),
  );
  sequencer.markTokenizationReady();
  sequencer.markAnnotationLoadingComplete('Subtitle annotations loaded');

  assert.deepEqual(osdMessages, ['Subtitle annotations loaded']);
});

test('startup OSD shows dictionary failure after annotation loading completes', () => {
  const osdMessages: string[] = [];
  const sequencer = createStartupOsdSequencer({
    showOsd: (message) => {
      osdMessages.push(message);
    },
  });

  sequencer.showAnnotationLoading('Loading subtitle annotations |');
  sequencer.notifyCharacterDictionaryStatus(
    makeDictionaryEvent('failed', 'Character dictionary sync failed for Frieren: boom'),
  );
  sequencer.markTokenizationReady();
  sequencer.markAnnotationLoadingComplete('Subtitle annotations loaded');

  assert.deepEqual(osdMessages, [
    'Loading subtitle annotations |',
    'Character dictionary sync failed for Frieren: boom',
  ]);
});

test('startup OSD reset keeps tokenization ready after first warmup', () => {
  const osdMessages: string[] = [];
  const sequencer = createStartupOsdSequencer({
    showOsd: (message) => {
      osdMessages.push(message);
    },
  });

  sequencer.markTokenizationReady();
  sequencer.reset();
  sequencer.notifyCharacterDictionaryStatus(
    makeDictionaryEvent('syncing', 'Updating character dictionary for Frieren...'),
  );

  assert.deepEqual(osdMessages, ['Updating character dictionary for Frieren...']);
});
