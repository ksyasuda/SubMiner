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

  assert.deepEqual(osdMessages, [
    'Loading subtitle annotations |',
    'Generating character dictionary for Frieren...',
  ]);

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

test('startup OSD shows dictionary ready when progress completed before it became visible', () => {
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

  assert.deepEqual(osdMessages, [
    'Character dictionary ready for Frieren',
    'Subtitle annotations loaded',
  ]);
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

test('startup OSD shows later dictionary progress immediately once tokenization is ready', () => {
  const osdMessages: string[] = [];
  const sequencer = createStartupOsdSequencer({
    showOsd: (message) => {
      osdMessages.push(message);
    },
  });

  sequencer.showAnnotationLoading('Loading subtitle annotations |');
  sequencer.markTokenizationReady();
  sequencer.notifyCharacterDictionaryStatus(
    makeDictionaryEvent('generating', 'Generating character dictionary for Frieren...'),
  );

  assert.deepEqual(osdMessages, [
    'Loading subtitle annotations |',
    'Generating character dictionary for Frieren...',
  ]);

  sequencer.markAnnotationLoadingComplete('Subtitle annotations loaded');

  assert.deepEqual(osdMessages, [
    'Loading subtitle annotations |',
    'Generating character dictionary for Frieren...',
  ]);
});

test('startup OSD keeps dictionary progress pending when mpv osd is unavailable', () => {
  const osdMessages: string[] = [];
  let osdAvailable = false;
  const sequencer = createStartupOsdSequencer({
    showOsd: (message) => {
      osdMessages.push(message);
      return osdAvailable;
    },
  });

  sequencer.markTokenizationReady();
  sequencer.notifyCharacterDictionaryStatus(
    makeDictionaryEvent('generating', 'Generating character dictionary for Frieren...'),
  );
  sequencer.notifyCharacterDictionaryStatus(
    makeDictionaryEvent('ready', 'Character dictionary ready for Frieren'),
  );

  assert.deepEqual(osdMessages, [
    'Generating character dictionary for Frieren...',
    'Character dictionary ready for Frieren',
  ]);

  osdAvailable = true;
  sequencer.notifyCharacterDictionaryStatus(
    makeDictionaryEvent('ready', 'Character dictionary ready for Frieren'),
  );

  assert.deepEqual(osdMessages, [
    'Generating character dictionary for Frieren...',
    'Character dictionary ready for Frieren',
    'Character dictionary ready for Frieren',
  ]);
});

test('startup notifications route tokenization and annotation status to overlay and system without osd for both', () => {
  const calls: string[] = [];
  const sequencer = createStartupOsdSequencer({
    getNotificationType: () => 'both',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showOverlayNotification: (payload) => {
      calls.push(
        `overlay:${payload.id}:${payload.title}:${payload.body}:${payload.variant}:${payload.persistent ? 'pin' : 'auto'}`,
      );
    },
    showDesktopNotification: (title, options) => {
      calls.push(`desktop:${title}:${options.body ?? ''}`);
    },
  });

  sequencer.showTokenizationLoading('Loading subtitle tokenization...');
  sequencer.markTokenizationReady();
  sequencer.showAnnotationLoading('Loading subtitle annotations |');
  sequencer.markAnnotationLoadingComplete('Subtitle annotations loaded');

  assert.deepEqual(calls, [
    'overlay:startup-tokenization:Subtitle tokenization:Loading subtitle tokenization...:progress:pin',
    'overlay:startup-tokenization:Subtitle tokenization:Subtitle tokenization ready:success:auto',
    'desktop:SubMiner:Subtitle tokenization ready',
    'overlay:startup-subtitle-annotations:Subtitle annotations:Loading subtitle annotations |:progress:pin',
    'overlay:startup-subtitle-annotations:Subtitle annotations:Subtitle annotations loaded:success:auto',
    'desktop:SubMiner:Subtitle annotations loaded',
  ]);
});
