const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const path = require('node:path');

const LINUX_FFMPEG_LIBRARY = 'libffmpeg.so';
const MACOS_WINDOW_HELPER = 'get-mpv-window-macos';
const MACOS_CONTROLLER_HELPER = 'macos-controller-watch';
const MACOS_HELPERS = [MACOS_WINDOW_HELPER, MACOS_CONTROLLER_HELPER];
const DEFAULT_MACOS_APP_BUNDLE = 'SubMiner.app';

function resolveMacOSAppBundlePath(context) {
  if (context.appOutDir.endsWith('.app')) {
    return context.appOutDir;
  }

  const productFilename = context.packager?.appInfo?.productFilename;
  const bundleName =
    typeof productFilename === 'string' && productFilename.trim()
      ? `${productFilename.trim()}.app`
      : DEFAULT_MACOS_APP_BUNDLE;

  return path.join(context.appOutDir, bundleName);
}

async function stageLinuxAppImageSharedLibrary(
  context,
  deps = {
    access: (filePath) => fs.access(filePath),
    mkdir: (dirPath) => fs.mkdir(dirPath, { recursive: true }),
    copyFile: (from, to) => fs.copyFile(from, to),
  },
) {
  if (context.electronPlatformName !== 'linux') return false;

  const sourceLibraryPath = path.join(context.appOutDir, LINUX_FFMPEG_LIBRARY);

  try {
    await deps.access(sourceLibraryPath);
  } catch (error) {
    if (error && typeof error === 'object' && error.code === 'ENOENT') {
      throw new Error(
        `Linux packaging requires ${LINUX_FFMPEG_LIBRARY} at ${sourceLibraryPath} so AppImage child processes can resolve it.`,
      );
    }
    throw error;
  }

  const targetLibraryDir = path.join(context.appOutDir, 'usr', 'lib');
  const targetLibraryPath = path.join(targetLibraryDir, LINUX_FFMPEG_LIBRARY);

  await deps.mkdir(targetLibraryDir);
  await deps.copyFile(sourceLibraryPath, targetLibraryPath);

  return true;
}

async function verifyMacOSHelpers(
  context,
  deps = {
    access: (filePath, mode) => fs.access(filePath, mode),
  },
) {
  if (context.electronPlatformName !== 'darwin') return false;

  const scriptsDir = path.join(
    resolveMacOSAppBundlePath(context),
    'Contents',
    'Resources',
    'scripts',
  );

  for (const helperName of MACOS_HELPERS) {
    const helperPath = path.join(scriptsDir, helperName);
    try {
      await deps.access(helperPath, fsSync.constants.X_OK);
    } catch (error) {
      if (error && typeof error === 'object' && error.code === 'ENOENT') {
        throw new Error(
          `macOS packaging requires ${helperName} at ${helperPath}. Run bun run build:assets on macOS with swiftc available before packaging.`,
        );
      }
      if (error && typeof error === 'object' && error.code === 'EACCES') {
        throw new Error(`macOS helper is not executable: ${helperPath}`);
      }
      throw error;
    }
  }

  return true;
}

async function stageBundledBunRuntime(context, deps = {}) {
  const stageBunRuntime =
    deps.stageBunRuntime ?? (await import('./stage-bun-runtime.mjs')).stageBunRuntime;
  const productFilename = context.packager?.appInfo?.productFilename;
  await stageBunRuntime({
    appOutDir: context.appOutDir,
    platform: context.electronPlatformName,
    arch: context.arch,
    productFilename:
      typeof productFilename === 'string' && productFilename.trim()
        ? productFilename.trim()
        : 'SubMiner',
  });
}

async function afterPack(context, deps = {}) {
  await stageLinuxAppImageSharedLibrary(context);
  await verifyMacOSHelpers(context);
  await stageBundledBunRuntime(context, deps);
  await (deps.auditPackage ?? require('./package-audit.cjs').auditPackage)(context);
}

module.exports = {
  LINUX_FFMPEG_LIBRARY,
  MACOS_HELPERS,
  resolveMacOSAppBundlePath,
  stageBundledBunRuntime,
  stageLinuxAppImageSharedLibrary,
  verifyMacOSHelpers,
  default: afterPack,
};
