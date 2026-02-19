import path from 'node:path';

type ExistsSync = (candidate: string) => boolean;

type ConfigPathOptions = {
  xdgConfigHome?: string;
  homeDir: string;
  existsSync: ExistsSync;
  appNames?: readonly string[];
  defaultAppName?: string;
};

const DEFAULT_APP_NAMES = ['SubMiner', 'subminer'] as const;
const DEFAULT_FILE_NAMES = ['config.jsonc', 'config.json'] as const;

export function resolveConfigBaseDirs(
  xdgConfigHome: string | undefined,
  homeDir: string,
): string[] {
  const fallbackBaseDir = path.join(homeDir, '.config');
  const primaryBaseDir = xdgConfigHome?.trim() || fallbackBaseDir;
  return Array.from(new Set([primaryBaseDir, fallbackBaseDir]));
}

function getAppNames(options: ConfigPathOptions): readonly string[] {
  return options.appNames ?? DEFAULT_APP_NAMES;
}

function getDefaultAppName(options: ConfigPathOptions): string {
  return options.defaultAppName ?? DEFAULT_APP_NAMES[0];
}

export function resolveConfigDir(options: ConfigPathOptions): string {
  const baseDirs = resolveConfigBaseDirs(options.xdgConfigHome, options.homeDir);
  const appNames = getAppNames(options);

  for (const baseDir of baseDirs) {
    for (const appName of appNames) {
      const dir = path.join(baseDir, appName);
      for (const fileName of DEFAULT_FILE_NAMES) {
        if (options.existsSync(path.join(dir, fileName))) {
          return dir;
        }
      }
    }
  }

  for (const baseDir of baseDirs) {
    for (const appName of appNames) {
      const dir = path.join(baseDir, appName);
      if (options.existsSync(dir)) {
        return dir;
      }
    }
  }

  return path.join(baseDirs[0], getDefaultAppName(options));
}

export function resolveConfigFilePath(options: ConfigPathOptions): string {
  const baseDirs = resolveConfigBaseDirs(options.xdgConfigHome, options.homeDir);
  const appNames = getAppNames(options);

  for (const baseDir of baseDirs) {
    for (const appName of appNames) {
      for (const fileName of DEFAULT_FILE_NAMES) {
        const candidate = path.join(baseDir, appName, fileName);
        if (options.existsSync(candidate)) {
          return candidate;
        }
      }
    }
  }

  return path.join(baseDirs[0], getDefaultAppName(options), DEFAULT_FILE_NAMES[0]);
}
