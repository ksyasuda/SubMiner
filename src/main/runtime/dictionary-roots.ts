// Candidate root directories searched for a bundled dictionary (JLPT vocab, frequency, ...).
// `vendorDirName` is the dictionary's folder under `vendor/` in dev and in packaged resources.
export function createBuildDictionaryRootsHandler(deps: {
  vendorDirName: string;
  platform: NodeJS.Platform;
  dirname: string;
  appPath: string;
  resourcesPath: string;
  userDataPath: string;
  appUserDataPath: string;
  homeDir: string;
  appDataDir?: string;
  cwd: string;
  joinPath: (...parts: string[]) => string;
}) {
  return (): string[] => {
    const platformRoots =
      deps.platform === 'win32'
        ? [
            deps.joinPath(
              deps.appDataDir ?? deps.joinPath(deps.homeDir, 'AppData', 'Roaming'),
              'SubMiner',
            ),
            deps.joinPath(
              deps.appDataDir ?? deps.joinPath(deps.homeDir, 'AppData', 'Roaming'),
              'subminer',
            ),
          ]
        : [
            deps.joinPath(deps.homeDir, '.config', 'SubMiner'),
            deps.joinPath(deps.homeDir, '.config', 'subminer'),
            deps.joinPath(deps.homeDir, 'Library', 'Application Support', 'SubMiner'),
            deps.joinPath(deps.homeDir, 'Library', 'Application Support', 'subminer'),
          ];

    return [
      deps.joinPath(deps.dirname, '..', '..', 'vendor', deps.vendorDirName),
      deps.joinPath(deps.appPath, 'vendor', deps.vendorDirName),
      deps.joinPath(deps.resourcesPath, deps.vendorDirName),
      deps.joinPath(deps.resourcesPath, 'app.asar', 'vendor', deps.vendorDirName),
      deps.userDataPath,
      deps.appUserDataPath,
      ...platformRoots,
      deps.cwd,
    ];
  };
}
