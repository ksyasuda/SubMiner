import * as fs from 'node:fs';
import * as path from 'node:path';

export interface YomitanExtensionPathOptions {
  explicitPath?: string;
  cwd?: string;
  moduleDir?: string;
  resourcesPath?: string;
  userDataPath?: string;
}

function pushUnique(values: string[], candidate: string | null | undefined): void {
  if (!candidate || values.includes(candidate)) {
    return;
  }
  values.push(candidate);
}

export function getYomitanExtensionSearchPaths(
  options: YomitanExtensionPathOptions = {},
): string[] {
  const searchPaths: string[] = [];

  pushUnique(searchPaths, options.explicitPath ? path.resolve(options.explicitPath) : null);
  pushUnique(searchPaths, options.cwd ? path.resolve(options.cwd, 'build', 'yomitan') : null);
  pushUnique(
    searchPaths,
    options.moduleDir
      ? path.resolve(options.moduleDir, '..', '..', '..', 'build', 'yomitan')
      : null,
  );
  pushUnique(
    searchPaths,
    options.resourcesPath ? path.join(options.resourcesPath, 'yomitan') : null,
  );
  pushUnique(searchPaths, '/usr/share/SubMiner/yomitan');
  pushUnique(searchPaths, options.userDataPath ? path.join(options.userDataPath, 'yomitan') : null);

  return searchPaths;
}

export function resolveExistingYomitanExtensionPath(
  searchPaths: string[],
  existsSync: (path: string) => boolean = fs.existsSync,
): string | null {
  for (const candidate of searchPaths) {
    if (existsSync(path.join(candidate, 'manifest.json'))) {
      return candidate;
    }
  }

  return null;
}

export function resolveYomitanExtensionPath(
  options: YomitanExtensionPathOptions = {},
  existsSync: (path: string) => boolean = fs.existsSync,
): string | null {
  return resolveExistingYomitanExtensionPath(getYomitanExtensionSearchPaths(options), existsSync);
}

export function resolveExternalYomitanExtensionPath(
  externalProfilePath: string,
  existsSync: (path: string) => boolean = fs.existsSync,
): string | null {
  const normalizedProfilePath = externalProfilePath.trim();
  if (!normalizedProfilePath) {
    return null;
  }

  const candidate = path.join(path.resolve(normalizedProfilePath), 'extensions', 'yomitan');
  return existsSync(path.join(candidate, 'manifest.json')) ? candidate : null;
}
