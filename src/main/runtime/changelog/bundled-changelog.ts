export function resolveBundledChangelogPath(deps: {
  resourcesPath: string;
  appPath: string;
  dirname: string;
  joinPath: (...parts: string[]) => string;
  fileExists: (path: string) => boolean;
}): string | null {
  const candidates = [
    deps.joinPath(deps.resourcesPath, 'CHANGELOG.md'),
    deps.joinPath(deps.appPath, 'CHANGELOG.md'),
    deps.joinPath(deps.dirname, '..', 'CHANGELOG.md'),
    deps.joinPath(deps.dirname, '..', '..', 'CHANGELOG.md'),
  ];

  return candidates.find((candidate) => deps.fileExists(candidate)) ?? null;
}

export function readBundledChangelog(deps: {
  resolvePath: () => string | null;
  readFile: (path: string) => string;
  logWarn: (message: string) => void;
}): string | null {
  const changelogPath = deps.resolvePath();
  if (!changelogPath) return null;
  try {
    return deps.readFile(changelogPath);
  } catch (error) {
    deps.logWarn(
      `Failed to read bundled changelog at ${changelogPath}: ${
        error instanceof Error ? error.message : String(error)
      }`,
    );
    return null;
  }
}
