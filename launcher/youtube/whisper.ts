import fs from 'node:fs';

import type { Args } from '../types.js';
import { runExternalCommand } from '../util.js';

export interface BuildWhisperArgsOptions {
  modelPath: string;
  audioPath: string;
  outputPrefix: string;
  language: string;
  translate: boolean;
  threads: number;
  vadModelPath?: string;
}

export function buildWhisperArgs(options: BuildWhisperArgsOptions): string[] {
  const args = [
    '-m',
    options.modelPath,
    '-f',
    options.audioPath,
    '--output-srt',
    '--output-file',
    options.outputPrefix,
    '--language',
    options.language,
    '--threads',
    String(options.threads),
  ];
  if (options.translate) args.push('--translate');
  if (options.vadModelPath) {
    args.push('-vm', options.vadModelPath, '--vad');
  }
  return args;
}

export async function runWhisper(
  whisperBin: string,
  args: Args,
  options: Omit<BuildWhisperArgsOptions, 'threads' | 'vadModelPath'>,
): Promise<string> {
  const vadModelPath =
    args.whisperVadModel.trim() && fs.existsSync(args.whisperVadModel.trim())
      ? args.whisperVadModel.trim()
      : undefined;
  const whisperArgs = buildWhisperArgs({
    ...options,
    threads: args.whisperThreads,
    vadModelPath,
  });
  await runExternalCommand(whisperBin, whisperArgs, {
    commandLabel: 'whisper',
    streamOutput: true,
  });
  const outputPath = `${options.outputPrefix}.srt`;
  if (!fs.existsSync(outputPath)) {
    throw new Error(`whisper output not found: ${outputPath}`);
  }
  return outputPath;
}
