import { access, stat } from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import type {
  SubtitleGenerationConfig,
  SubtitleGenerationToolStatus,
  SubtitleGenerationTools,
} from '../../shared/subtitle-generation';
import { expandSubtitleGenerationPath } from './subtitle-generation-files';

/** Executable paths ready to spawn. `vad` is null when dialogue mode is off. */
export interface SubtitleGenerationToolPaths {
  ffmpeg: string;
  ffprobe: string;
  whisper: string;
  vad: string | null;
}

const TOOL_LOOKUPS = {
  ffmpeg: { setting: 'ffmpegPath', names: ['ffmpeg'], install: 'Install FFmpeg' },
  ffprobe: { setting: 'ffprobePath', names: ['ffprobe'], install: 'Install FFmpeg' },
  whisper: { setting: 'whisperPath', names: ['whisper-cli'], install: 'Install whisper.cpp' },
  vad: {
    setting: 'vadPath',
    names: ['whisper-vad-speech-segments', 'vad-speech-segments'],
    install: "Install whisper.cpp's speech segment detector",
  },
} as const;

async function isExecutableFile(filePath: string): Promise<boolean> {
  try {
    if (!(await stat(filePath)).isFile()) return false;
    await access(filePath, constants.X_OK);
    return true;
  } catch {
    return false;
  }
}

function executableNames(name: string, env: NodeJS.ProcessEnv): string[] {
  if (process.platform !== 'win32' || path.extname(name)) return [name];
  const extensions = (env.PATHEXT ?? '.EXE;.CMD;.BAT')
    .split(';')
    .map((entry) => entry.trim())
    .filter(Boolean);
  return [name, ...extensions.map((extension) => `${name}${extension}`)];
}

async function findOnPath(names: readonly string[], env: NodeJS.ProcessEnv): Promise<string> {
  const directories = (env.PATH ?? '')
    .split(path.delimiter)
    .map((entry) => entry.trim())
    .filter(Boolean);
  for (const directory of directories) {
    for (const name of names) {
      for (const candidate of executableNames(name, env)) {
        const filePath = path.join(directory, candidate);
        if (await isExecutableFile(filePath)) return filePath;
      }
    }
  }
  return '';
}

async function resolveTool(
  tool: keyof typeof TOOL_LOOKUPS,
  config: SubtitleGenerationConfig,
  env: NodeJS.ProcessEnv,
): Promise<SubtitleGenerationToolStatus> {
  const lookup = TOOL_LOOKUPS[tool];
  const override = config[lookup.setting].trim();
  if (override) {
    const expanded = expandSubtitleGenerationPath(override);
    const found =
      path.dirname(expanded) === '.'
        ? await findOnPath([expanded], env)
        : (await isExecutableFile(expanded))
          ? path.resolve(expanded)
          : '';
    return found
      ? { kind: 'found', path: found }
      : {
          kind: 'missing',
          message: `${override} (subtitleGeneration.${lookup.setting}) is not an executable file.`,
        };
  }
  const found = await findOnPath(lookup.names, env);
  return found
    ? { kind: 'found', path: found }
    : {
        kind: 'missing',
        message: `${lookup.names[0]} was not found on PATH. ${lookup.install} or set subtitleGeneration.${lookup.setting} in Settings.`,
      };
}

/** Locate every executable a generation run needs, before any model download or audio work. */
export async function resolveSubtitleGenerationTools(
  config: SubtitleGenerationConfig,
  env: NodeJS.ProcessEnv = process.env,
): Promise<SubtitleGenerationTools> {
  const [ffmpeg, ffprobe, whisper, vad] = await Promise.all([
    resolveTool('ffmpeg', config, env),
    resolveTool('ffprobe', config, env),
    resolveTool('whisper', config, env),
    config.vadModelPath.trim() ? resolveTool('vad', config, env) : null,
  ]);
  return { ffmpeg, ffprobe, whisper, vad };
}

function foundPath(tool: SubtitleGenerationToolStatus): string {
  if (tool.kind === 'missing') throw new Error(tool.message);
  return tool.path;
}

/** Throw the first missing tool's message, otherwise narrow to spawnable paths. */
export function requireSubtitleGenerationTools(
  tools: SubtitleGenerationTools,
): SubtitleGenerationToolPaths {
  return {
    ffmpeg: foundPath(tools.ffmpeg),
    ffprobe: foundPath(tools.ffprobe),
    whisper: foundPath(tools.whisper),
    vad: tools.vad ? foundPath(tools.vad) : null,
  };
}
