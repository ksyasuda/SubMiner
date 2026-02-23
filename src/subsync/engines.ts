export type SubsyncEngine = 'alass' | 'ffsubsync';

export interface SubsyncCommandResult {
  ok: boolean;
  code: number | null;
  stderr: string;
  stdout: string;
  error?: string;
}

export interface SubsyncEngineExecutionContext {
  referenceFilePath: string;
  videoPath: string;
  inputSubtitlePath: string;
  outputPath: string;
  audioStreamIndex: number | null;
  resolveExecutablePath: (configuredPath: string, commandName: string) => string;
  resolvedPaths: {
    alassPath: string;
    ffsubsyncPath: string;
  };
  runCommand: (command: string, args: string[]) => Promise<SubsyncCommandResult>;
}

export interface SubsyncEngineProvider {
  engine: SubsyncEngine;
  execute: (context: SubsyncEngineExecutionContext) => Promise<SubsyncCommandResult>;
}

type SubsyncEngineProviderFactory = () => SubsyncEngineProvider;

const subsyncEngineProviderFactories = new Map<SubsyncEngine, SubsyncEngineProviderFactory>();

export function registerSubsyncEngineProvider(
  engine: SubsyncEngine,
  factory: SubsyncEngineProviderFactory,
): void {
  if (subsyncEngineProviderFactories.has(engine)) {
    return;
  }
  subsyncEngineProviderFactories.set(engine, factory);
}

export function createSubsyncEngineProvider(engine: SubsyncEngine): SubsyncEngineProvider | null {
  const factory = subsyncEngineProviderFactories.get(engine);
  if (!factory) return null;
  return factory();
}

function registerDefaultSubsyncEngineProviders(): void {
  registerSubsyncEngineProvider('alass', () => ({
    engine: 'alass',
    execute: async (context: SubsyncEngineExecutionContext) => {
      const alassPath = context.resolveExecutablePath(context.resolvedPaths.alassPath, 'alass');
      return context.runCommand(alassPath, [
        context.referenceFilePath,
        context.inputSubtitlePath,
        context.outputPath,
      ]);
    },
  }));

  registerSubsyncEngineProvider('ffsubsync', () => ({
    engine: 'ffsubsync',
    execute: async (context: SubsyncEngineExecutionContext) => {
      const ffsubsyncPath = context.resolveExecutablePath(
        context.resolvedPaths.ffsubsyncPath,
        'ffsubsync',
      );
      const args = [context.videoPath, '-i', context.inputSubtitlePath, '-o', context.outputPath];
      if (context.audioStreamIndex !== null) {
        args.push('--reference-stream', `0:${context.audioStreamIndex}`);
      }
      return context.runCommand(ffsubsyncPath, args);
    },
  }));
}

registerDefaultSubsyncEngineProviders();
