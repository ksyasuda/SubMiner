import fs from 'node:fs';
import path from 'node:path';
import type { MpvInstallPaths } from '../../shared/setup-state';

export interface InstalledFirstRunPluginCandidate {
  path: string;
  kind: 'directory' | 'file';
}

export type InstalledMpvPluginSource =
  | 'default-config'
  | 'xdg-config'
  | 'portable-config'
  | 'legacy-file';

export interface InstalledMpvPluginDetection {
  installed: boolean;
  path: string | null;
  version: string | null;
  source: InstalledMpvPluginSource | null;
  message: string | null;
}

export interface LegacyMpvPluginRemovalResult {
  ok: boolean;
  removedPaths: string[];
  failedPaths: Array<{ path: string; message: string }>;
}

export function resolvePackagedFirstRunPluginAssets(deps: {
  dirname: string;
  appPath: string;
  resourcesPath: string;
  joinPath?: (...parts: string[]) => string;
  existsSync?: (candidate: string) => boolean;
}): { pluginDirSource: string; pluginConfigSource: string } | null {
  const joinPath = deps.joinPath ?? path.join;
  const existsSync = deps.existsSync ?? fs.existsSync;
  const roots = [
    joinPath(deps.resourcesPath, 'plugin'),
    joinPath(deps.resourcesPath, 'app.asar', 'plugin'),
    joinPath(deps.appPath, 'plugin'),
    joinPath(deps.dirname, '..', 'plugin'),
    joinPath(deps.dirname, '..', '..', 'plugin'),
  ];

  for (const root of roots) {
    const pluginDirSource = joinPath(root, 'subminer');
    const pluginConfigSource = joinPath(root, 'subminer.conf');
    if (
      existsSync(pluginDirSource) &&
      existsSync(pluginConfigSource) &&
      existsSync(joinPath(pluginDirSource, 'main.lua'))
    ) {
      return { pluginDirSource, pluginConfigSource };
    }
  }

  return null;
}

export function resolvePackagedRuntimePluginPath(deps: {
  dirname: string;
  appPath: string;
  resourcesPath: string;
  joinPath?: (...parts: string[]) => string;
  existsSync?: (candidate: string) => boolean;
}): string | null {
  const joinPath = deps.joinPath ?? path.join;
  const existsSync = deps.existsSync ?? fs.existsSync;
  const assets = resolvePackagedFirstRunPluginAssets({
    dirname: deps.dirname,
    appPath: deps.appPath,
    resourcesPath: deps.resourcesPath,
    joinPath,
    existsSync,
  });
  if (!assets) {
    return null;
  }

  const entrypoint = joinPath(assets.pluginDirSource, 'main.lua');
  return existsSync(entrypoint) ? entrypoint : null;
}

export function detectInstalledFirstRunPlugin(
  installPaths: MpvInstallPaths,
  deps?: {
    existsSync?: (candidate: string) => boolean;
  },
): boolean {
  const existsSync = deps?.existsSync ?? fs.existsSync;
  const pluginEntrypointPath = path.join(installPaths.scriptsDir, 'subminer', 'main.lua');
  return existsSync(pluginEntrypointPath);
}

function getPlatformPath(platform: NodeJS.Platform): typeof path.posix | typeof path.win32 {
  return platform === 'win32' ? path.win32 : path.posix;
}

interface MpvConfigRootCandidate {
  root: string;
  source: Exclude<InstalledMpvPluginSource, 'legacy-file'>;
}

function collectMpvConfigRootCandidates(options: {
  platform: NodeJS.Platform;
  homeDir: string;
  xdgConfigHome?: string;
  appDataDir?: string;
  mpvExecutablePath?: string;
}): MpvConfigRootCandidate[] {
  const platformPath = getPlatformPath(options.platform);
  if (options.platform === 'win32') {
    const roots: MpvConfigRootCandidate[] = [];
    if (options.mpvExecutablePath?.trim()) {
      roots.push({
        root: platformPath.join(
          platformPath.dirname(options.mpvExecutablePath.trim()),
          'portable_config',
        ),
        source: 'portable-config',
      });
    }
    roots.push({
      root: platformPath.join(
        options.appDataDir?.trim() || platformPath.join(options.homeDir, 'AppData', 'Roaming'),
        'mpv',
      ),
      source: 'default-config',
    });
    return roots;
  }

  const xdgRoot = options.xdgConfigHome?.trim()
    ? platformPath.join(options.xdgConfigHome.trim(), 'mpv')
    : null;
  const homeRoot = platformPath.join(options.homeDir, '.config', 'mpv');
  const roots: MpvConfigRootCandidate[] = [];
  if (xdgRoot) {
    roots.push({ root: xdgRoot, source: 'xdg-config' });
  }
  if (!xdgRoot || xdgRoot !== homeRoot) {
    roots.push({ root: homeRoot, source: 'default-config' });
  }
  return roots;
}

export function detectInstalledFirstRunPluginCandidates(options: {
  platform: NodeJS.Platform;
  homeDir: string;
  xdgConfigHome?: string;
  appDataDir?: string;
  mpvExecutablePath?: string;
  existsSync?: (candidate: string) => boolean;
}): InstalledFirstRunPluginCandidate[] {
  const platformPath = getPlatformPath(options.platform);
  const existsSync = options.existsSync ?? fs.existsSync;
  const roots = collectMpvConfigRootCandidates(options);

  const candidates: InstalledFirstRunPluginCandidate[] = [];
  const seen = new Set<string>();
  const pushIfExists = (
    candidate: InstalledFirstRunPluginCandidate,
    verifyPath = candidate.path,
  ) => {
    if (seen.has(candidate.path) || !existsSync(verifyPath)) return;
    seen.add(candidate.path);
    candidates.push(candidate);
  };

  for (const root of roots) {
    const scriptsDir = platformPath.join(root.root, 'scripts');
    const pluginDir = platformPath.join(scriptsDir, 'subminer');
    pushIfExists({ path: pluginDir, kind: 'directory' });
    pushIfExists({ path: platformPath.join(scriptsDir, 'subminer.lua'), kind: 'file' });
    pushIfExists({ path: platformPath.join(scriptsDir, 'subminer-loader.lua'), kind: 'file' });
  }

  return candidates;
}

function parseInstalledPluginVersion(content: string): string | null {
  const match = content.match(/\bversion\s*=\s*["']([^"']+)["']/);
  return match?.[1] ?? null;
}

function readInstalledPluginVersion(options: {
  pluginEntrypointPath: string;
  platformPath: typeof path.posix | typeof path.win32;
  existsSync: (candidate: string) => boolean;
  readFileSync: (candidate: string, encoding: BufferEncoding) => string;
}): string | null {
  const versionPath = options.platformPath.join(
    options.platformPath.dirname(options.pluginEntrypointPath),
    'version.lua',
  );
  if (!options.existsSync(versionPath)) {
    return null;
  }
  try {
    return parseInstalledPluginVersion(options.readFileSync(versionPath, 'utf8'));
  } catch {
    return null;
  }
}

export function detectInstalledMpvPlugin(options: {
  platform: NodeJS.Platform;
  homeDir: string;
  xdgConfigHome?: string;
  appDataDir?: string;
  mpvExecutablePath?: string;
  existsSync?: (candidate: string) => boolean;
  readFileSync?: (candidate: string, encoding: BufferEncoding) => string;
}): InstalledMpvPluginDetection {
  const platformPath = getPlatformPath(options.platform);
  const existsSync = options.existsSync ?? fs.existsSync;
  const readFileSync =
    options.readFileSync ?? ((candidate, encoding) => fs.readFileSync(candidate, encoding));
  const roots = collectMpvConfigRootCandidates(options);

  for (const root of roots) {
    const scriptsDir = platformPath.join(root.root, 'scripts');
    const directoryEntrypoint = platformPath.join(scriptsDir, 'subminer', 'main.lua');
    if (existsSync(directoryEntrypoint)) {
      const version = readInstalledPluginVersion({
        pluginEntrypointPath: directoryEntrypoint,
        platformPath,
        existsSync,
        readFileSync,
      });
      return {
        installed: true,
        path: directoryEntrypoint,
        version,
        source: root.source,
        message: `SubMiner detected an installed mpv plugin at: ${directoryEntrypoint}`,
      };
    }

    for (const legacyPath of [
      platformPath.join(scriptsDir, 'subminer.lua'),
      platformPath.join(scriptsDir, 'subminer-loader.lua'),
    ]) {
      if (existsSync(legacyPath)) {
        return {
          installed: true,
          path: legacyPath,
          version: null,
          source: root.source === 'portable-config' ? 'portable-config' : 'legacy-file',
          message: `SubMiner detected an installed mpv plugin at: ${legacyPath}`,
        };
      }
    }
  }

  return {
    installed: false,
    path: null,
    version: null,
    source: null,
    message: null,
  };
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export async function removeLegacyMpvPluginCandidates(options: {
  candidates: InstalledFirstRunPluginCandidate[];
  trashItem: (path: string) => Promise<void>;
}): Promise<LegacyMpvPluginRemovalResult> {
  const removedPaths: string[] = [];
  const failedPaths: Array<{ path: string; message: string }> = [];
  const seen = new Set<string>();

  for (const candidate of options.candidates) {
    if (seen.has(candidate.path)) continue;
    seen.add(candidate.path);
    try {
      await options.trashItem(candidate.path);
      removedPaths.push(candidate.path);
    } catch (error) {
      failedPaths.push({ path: candidate.path, message: errorMessage(error) });
    }
  }

  return {
    ok: failedPaths.length === 0,
    removedPaths,
    failedPaths,
  };
}
