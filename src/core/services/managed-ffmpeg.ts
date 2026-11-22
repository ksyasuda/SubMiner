import { execFile } from 'node:child_process';
import { createReadStream, createWriteStream, existsSync } from 'node:fs';
import { chmod, mkdir, mkdtemp, rename, rm } from 'node:fs/promises';
import path from 'node:path';
import { pipeline } from 'node:stream/promises';
import { createGunzip } from 'node:zlib';
import { downloadSubtitleGenerationArtifact } from './subtitle-generation-download';
import { findOnPath } from './subtitle-generation-tools';

/**
 * SubMiner can download its own FFmpeg when none is usable on PATH. The copy lives in
 * `<configDir>/tools/ffmpeg` and is appended to PATH at startup, so a system FFmpeg always
 * wins and every bare `ffmpeg`/`ffprobe` spawn picks up the download otherwise.
 */

export type FfmpegTool = 'ffmpeg' | 'ffprobe';

const FFMPEG_TOOLS: readonly FfmpegTool[] = ['ffmpeg', 'ffprobe'];

/** Pinned ffmpeg-static release: single gzipped static binaries per platform. */
const MANAGED_FFMPEG_RELEASE_URL =
  'https://github.com/eugeneware/ffmpeg-static/releases/download/b6.1.1';

type ManagedAsset = { size: number; sha256: string };

const MANAGED_FFMPEG_ASSETS = {
  'linux-x64': {
    ffmpeg: {
      size: 29354986,
      sha256: 'bfe8a8fc511530457b528c48d77b5737527b504a3797a9bc4866aeca69c2dffa',
    },
    ffprobe: {
      size: 29276839,
      sha256: '25d9b6ccb05e3d9de9e04e31e2506d8dd7f9f0418981965ac6df12e8d3afd067',
    },
  },
  'linux-arm64': {
    ffmpeg: {
      size: 25568691,
      sha256: '754a678672298bc68156adff58aa7385a592c2b30b1d0ae8750c45c915c4bac0',
    },
    ffprobe: {
      size: 25493573,
      sha256: '2ab6aba60ee84412dff9188720703376cb4e7aaf7e0b5e43aa8249f2acae5bf8',
    },
  },
  'darwin-x64': {
    ffmpeg: {
      size: 25296431,
      sha256: '929b375c1182d956c51f7ac25e0b2b0411fb01f6f407aa15c9758efeb4242106',
    },
    ffprobe: {
      size: 25239438,
      sha256: 'd4da574d6e2e197bd259b47d69cf262df9e312af24ad960444f6d806d3d4c186',
    },
  },
  'darwin-arm64': {
    ffmpeg: {
      size: 19246198,
      sha256: '8923876afa8db5585022d7860ec7e589af192f441c56793971276d450ed3bbfa',
    },
    ffprobe: {
      size: 19207077,
      sha256: 'd986a8ec7b030899fe66a8a288ed809a3543338705a3ce178cfb85869c5d80be',
    },
  },
  'win32-x64': {
    ffmpeg: {
      size: 29581307,
      sha256: '8883a3dffbd0a16cf4ef95206ea05283f78908dbfb118f73c83f4951dcc06d77',
    },
    ffprobe: {
      size: 29521644,
      sha256: 'f309e6223ad89d2fe54bccd420a7709b66fd27540674e92309578ed491a43c8d',
    },
  },
} as const satisfies Record<string, Record<FfmpegTool, ManagedAsset>>;

type ManagedFfmpegTarget = keyof typeof MANAGED_FFMPEG_ASSETS;

function getManagedFfmpegTarget(
  platform: NodeJS.Platform,
  arch: string,
): ManagedFfmpegTarget | null {
  // Windows on ARM runs the x64 build through emulation; there is no native arm64 build.
  const key = platform === 'win32' && arch === 'arm64' ? 'win32-x64' : `${platform}-${arch}`;
  return key in MANAGED_FFMPEG_ASSETS ? (key as ManagedFfmpegTarget) : null;
}

export function getManagedFfmpegDir(configDir: string): string {
  return path.join(configDir, 'tools', 'ffmpeg');
}

function managedBinaryPath(
  managedDir: string,
  tool: FfmpegTool,
  platform: NodeJS.Platform,
): string {
  return path.join(managedDir, platform === 'win32' ? `${tool}.exe` : tool);
}

function pathKey(env: NodeJS.ProcessEnv): string {
  return Object.keys(env).find((key) => key.toLowerCase() === 'path') ?? 'PATH';
}

/**
 * Appends the managed FFmpeg directory to `env`'s PATH when a download exists there.
 * Appending keeps a system FFmpeg in front. Returns whether PATH changed.
 */
export function appendManagedFfmpegToPath(
  env: NodeJS.ProcessEnv,
  managedDir: string,
  platform: NodeJS.Platform = process.platform,
): boolean {
  if (!existsSync(managedBinaryPath(managedDir, 'ffmpeg', platform))) return false;
  const key = pathKey(env);
  const entries = (env[key] ?? '').split(path.delimiter).filter(Boolean);
  if (entries.some((entry) => path.resolve(entry) === path.resolve(managedDir))) return false;
  env[key] = [...entries, managedDir].join(path.delimiter);
  return true;
}

export type FfmpegSetupStatus =
  | { kind: 'ready'; source: 'system' | 'managed'; ffmpegPath: string; ffprobePath: string }
  | { kind: 'missing'; missing: FfmpegTool[]; downloadable: boolean }
  /** Found on PATH but `-version` failed. Only a broken managed copy is fixed by downloading. */
  | { kind: 'broken'; path: string; message: string; managed: boolean; downloadable: boolean };

export type RunVersionCheck = (executablePath: string) => Promise<void>;

const runVersionCheckDefault: RunVersionCheck = (executablePath) =>
  new Promise((resolve, reject) => {
    // Generous timeout: Windows Defender scans a fresh 80 MB exe on first launch.
    execFile(executablePath, ['-version'], { timeout: 15_000 }, (error) =>
      error ? reject(error) : resolve(),
    );
  });

/** Whether SubMiner can run FFmpeg and FFprobe from the current PATH, and where they come from. */
export async function detectFfmpeg(options: {
  managedDir: string;
  env?: NodeJS.ProcessEnv;
  platform?: NodeJS.Platform;
  arch?: string;
  runVersionCheck?: RunVersionCheck;
}): Promise<FfmpegSetupStatus> {
  const env = options.env ?? process.env;
  const platform = options.platform ?? process.platform;
  const runVersionCheck = options.runVersionCheck ?? runVersionCheckDefault;
  const downloadable = getManagedFfmpegTarget(platform, options.arch ?? process.arch) !== null;
  const isManaged = (filePath: string) =>
    path.resolve(path.dirname(filePath)) === path.resolve(options.managedDir);

  const found = await Promise.all(FFMPEG_TOOLS.map((tool) => findOnPath([tool], env)));
  // Check what was found before reporting gaps: a broken system tool shadows any download.
  for (const executablePath of found.filter(Boolean)) {
    try {
      await runVersionCheck(executablePath);
    } catch (error) {
      return {
        kind: 'broken',
        path: executablePath,
        message: error instanceof Error ? error.message : String(error),
        managed: isManaged(executablePath),
        downloadable,
      };
    }
  }
  const missing = FFMPEG_TOOLS.filter((_, index) => !found[index]);
  if (missing.length > 0) return { kind: 'missing', missing, downloadable };

  const [ffmpegPath, ffprobePath] = found as [string, string];
  return {
    kind: 'ready',
    source: isManaged(ffmpegPath) || isManaged(ffprobePath) ? 'managed' : 'system',
    ffmpegPath,
    ffprobePath,
  };
}

/**
 * Downloads, verifies, and unpacks FFmpeg and FFprobe into `managedDir`, replacing any
 * earlier download. Callers append `managedDir` to PATH afterwards.
 */
export async function installManagedFfmpeg(options: {
  managedDir: string;
  platform?: NodeJS.Platform;
  arch?: string;
}): Promise<void> {
  const platform = options.platform ?? process.platform;
  const arch = options.arch ?? process.arch;
  const target = getManagedFfmpegTarget(platform, arch);
  if (!target) {
    throw new Error(
      `SubMiner has no FFmpeg download for ${platform}-${arch}. Install FFmpeg yourself.`,
    );
  }
  await mkdir(options.managedDir, { recursive: true });
  const staging = await mkdtemp(path.join(options.managedDir, '.install-'));
  try {
    const unpacked: Array<[string, string]> = [];
    for (const tool of FFMPEG_TOOLS) {
      const asset = MANAGED_FFMPEG_ASSETS[target][tool];
      const archivePath = path.join(staging, `${tool}.gz`);
      await downloadSubtitleGenerationArtifact({
        url: `${MANAGED_FFMPEG_RELEASE_URL}/${tool}-${target}.gz`,
        size: asset.size,
        sha256: asset.sha256,
        destination: archivePath,
        label: tool === 'ffmpeg' ? 'FFmpeg' : 'FFprobe',
      });
      const binaryPath = managedBinaryPath(staging, tool, platform);
      await pipeline(createReadStream(archivePath), createGunzip(), createWriteStream(binaryPath));
      await chmod(binaryPath, 0o755);
      unpacked.push([binaryPath, managedBinaryPath(options.managedDir, tool, platform)]);
    }
    // Publish only after both tools verified, so a failed download never leaves half a pair.
    for (const [from, to] of unpacked) await rename(from, to);
  } finally {
    await rm(staging, { recursive: true, force: true });
  }
}
