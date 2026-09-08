import assert from 'node:assert/strict';
import test from 'node:test';
import path from 'node:path';
import { parseArgs } from '../config.js';
import {
  createGenerationProgressReporter,
  runGenerateSubtitlesCommand,
} from './generate-subtitles-command.js';

type Deps = NonNullable<Parameters<typeof runGenerateSubtitlesCommand>[1]>;

function fixture(argv: string[] = ['generate-subs', '/media/episode.mkv']) {
  const output: string[] = [];
  const commands: unknown[][] = [];
  const generations: Parameters<NonNullable<Deps['generate']>>[0][] = [];
  let exitCode: number | undefined;
  let interrupted: (() => void) | undefined;
  let detached = false;
  const context = {
    args: parseArgs(argv, 'subminer', {}),
    mpvSocketPath: '/tmp/test-subminer-socket',
    processAdapter: {
      writeStdout: (text: string) => {
        output.push(text);
      },
      setExitCode: (code: number) => {
        exitCode = code;
      },
    },
  };
  const deps: Deps = {
    readConfig: () => ({ subtitleGeneration: { modelPath: '/models/external.bin' } }),
    configPath: () => '/settings/SubMiner/config.jsonc',
    resolveModel: async () => ({ kind: 'external', path: '/models/external.bin' }),
    downloadModel: async () => {
      throw new Error('Unexpected model download');
    },
    generate: async (input) => {
      generations.push(input);
      input.onProgress?.({
        stage: 'transcribe',
        percent: 50,
        message: 'Transcribing Japanese audio',
      });
      return '/media/episode.ja.srt';
    },
    mpvCommand: async (_socket, command) => {
      commands.push(command);
      if (command[1] === 'path') return '/media/episode.mkv';
      if (command[1] === 'track-list') return [{ type: 'audio', selected: true, 'ff-index': 2 }];
      return undefined;
    },
    onInterrupt: (handler) => {
      interrupted = handler;
      return () => {
        detached = true;
      };
    },
  };
  return {
    context,
    deps,
    output,
    commands,
    generations,
    exitCode: () => exitCode,
    interrupt: () => interrupted?.(),
    detached: () => detached,
  };
}

test('launcher uses the shared core and selected mpv audio then loads the generated file', async () => {
  const f = fixture(['generate-subs']);
  assert.equal(await runGenerateSubtitlesCommand(f.context, f.deps), true);
  assert.equal(f.generations[0]?.mediaPath, '/media/episode.mkv');
  assert.equal(f.generations[0]?.audioStreamIndex, 2);
  assert.equal(
    f.generations[0]?.modelDirectory,
    path.join('/settings/SubMiner', 'models', 'whisper'),
  );
  assert.equal(f.generations[0]?.config.modelPath, '/models/external.bin');
  assert.deepEqual(f.commands.at(-2), [
    'sub-add',
    '/media/episode.ja.srt',
    'select',
    'Japanese (generated)',
    'ja',
  ]);
  assert.deepEqual(f.commands.at(-1), ['set_property', 'sub-delay', 0]);
  assert.match(f.output.join(''), /50%/);
  assert.match(f.output.join(''), /Saved Japanese subtitles/);
  assert.equal(f.detached(), true);
});

test('launcher never downloads a model without the explicit option', async () => {
  const f = fixture();
  f.deps.resolveModel = async () => ({ kind: 'missing', path: '/models/missing.bin' });
  await assert.rejects(runGenerateSubtitlesCommand(f.context, f.deps), /--download-model/);
  assert.equal(f.generations.length, 0);
  assert.equal(f.detached(), true);
});

test('current mpv generation requires an identifiable selected audio track', async () => {
  for (const tracks of [
    [],
    [{ type: 'audio', selected: true }],
    [{ type: 'audio', selected: true, external: true, 'ff-index': 0 }],
  ]) {
    const f = fixture(['generate-subs']);
    f.deps.mpvCommand = async (_socket, command) =>
      command[1] === 'path' ? '/media/episode.mkv' : tracks;
    await assert.rejects(runGenerateSubtitlesCommand(f.context, f.deps), /audio track/);
    assert.equal(f.generations.length, 0);
  }
});

test('explicit local file leaves Japanese track selection to the shared generator', async () => {
  const f = fixture();
  await runGenerateSubtitlesCommand(f.context, f.deps);
  assert.equal(f.generations[0]?.audioStreamIndex, undefined);
  assert.equal(
    f.commands.some((command) => command[1] === 'track-list'),
    false,
  );
});

test('launcher does not load generated subtitles after mpv switches files', async () => {
  const f = fixture(['generate-subs']);
  let pathRequests = 0;
  f.deps.mpvCommand = async (_socket, command) => {
    f.commands.push(command);
    if (command[1] === 'path')
      return ++pathRequests === 1 ? '/media/episode.mkv' : '/media/next.mkv';
    return [{ type: 'audio', selected: true, 'ff-index': 2 }];
  };
  await runGenerateSubtitlesCommand(f.context, f.deps);
  assert.equal(
    f.commands.some((command) => command[0] === 'sub-add'),
    false,
  );
  assert.match(f.output.join(''), /Saved Japanese subtitles/);
});

test('explicit managed model overrides external config and downloads before generation', async () => {
  const f = fixture([
    'generate-subs',
    '/media/episode.mkv',
    '--model',
    'medium',
    '--download-model',
  ]);
  f.deps.resolveModel = async (config) => {
    assert.equal(config.modelPath, '');
    assert.equal(config.managedModel, 'medium');
    return { kind: 'missing', path: '/models/medium.bin' };
  };
  let downloaded = false;
  f.deps.downloadModel = async () => {
    downloaded = true;
    return '/models/medium.bin';
  };
  const generate = f.deps.generate;
  f.deps.generate = async (input) => {
    assert.equal(downloaded, true);
    if (!generate) throw new Error('Missing fixture generator');
    return generate(input);
  };
  await runGenerateSubtitlesCommand(f.context, f.deps);
  assert.equal(f.generations.length, 1);
});

test('generation can run standalone and never loads subtitles into another video', async () => {
  for (const playing of [null, '/media/different.mkv']) {
    const f = fixture();
    f.deps.mpvCommand = async (_socket, command) => {
      f.commands.push(command);
      if (playing === null) throw new Error('mpv is not running');
      return playing;
    };
    await runGenerateSubtitlesCommand(f.context, f.deps);
    assert.equal(f.generations.length, 1);
    assert.equal(
      f.commands.some((command) => command[0] === 'sub-add'),
      false,
    );
  }
});

test('launcher preserves the saved path when loading into mpv fails', async () => {
  const f = fixture();
  const mpv = f.deps.mpvCommand;
  f.deps.mpvCommand = async (socket, command, timeout) => {
    if (command[0] === 'sub-add') throw new Error('load failed');
    return mpv?.(socket, command, timeout);
  };
  await runGenerateSubtitlesCommand(f.context, f.deps);
  assert.match(f.output.join(''), /Saved Japanese subtitles: \/media\/episode.ja.srt/);
  assert.match(f.output.join(''), /mpv could not load them: load failed/);
  assert.equal(f.exitCode(), 1);
});

test('SIGINT cancels shared generation and unregisters its handler', async () => {
  const f = fixture();
  f.deps.generate = async (input) => {
    f.interrupt();
    assert.equal(input.signal?.aborted, true);
    throw new Error('Aborted');
  };
  await runGenerateSubtitlesCommand(f.context, f.deps);
  assert.equal(f.exitCode(), 130);
  assert.equal(f.detached(), true);
  assert.match(f.output.join(''), /cancelled/);
});

test('cancellation after generation preserves the saved path and skips mpv loading', async () => {
  const f = fixture();
  f.deps.generate = async () => {
    f.interrupt();
    return '/media/episode.ja.srt';
  };
  await runGenerateSubtitlesCommand(f.context, f.deps);
  assert.equal(
    f.commands.some((command) => command[0] === 'sub-add'),
    false,
  );
  assert.match(f.output.join(''), /Saved Japanese subtitles: \/media\/episode.ja.srt/);
  assert.equal(f.exitCode(), 130);
  assert.equal(f.detached(), true);
});

test('progress throttles repeated updates but always reports stage changes and completion', () => {
  const output: string[] = [];
  let time = 0;
  const progress = createGenerationProgressReporter(
    (text) => output.push(text),
    () => time,
  );
  progress({ stage: 'download', percent: 0, message: 'Downloading' });
  progress({ stage: 'download', percent: 1, message: 'Downloading' });
  time = 1000;
  progress({ stage: 'download', percent: 50, message: 'Downloading' });
  progress({ stage: 'download', percent: 100, message: 'Downloading' });
  progress({ stage: 'extract', message: 'Extracting audio' });
  assert.equal(output.length, 4);
});
