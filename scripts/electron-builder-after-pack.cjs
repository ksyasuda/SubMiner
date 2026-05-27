const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const path = require('node:path');

const LINUX_FFMPEG_LIBRARY = 'libffmpeg.so';
const MACOS_WINDOW_HELPER = 'get-mpv-window-macos';
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

async function verifyMacOSWindowHelper(
  context,
  deps = {
    access: (filePath, mode) => fs.access(filePath, mode),
  },
) {
  if (context.electronPlatformName !== 'darwin') return false;

  const appBundlePath = resolveMacOSAppBundlePath(context);
  const helperPath = path.join(
    appBundlePath,
    'Contents',
    'Resources',
    'scripts',
    MACOS_WINDOW_HELPER,
  );

  try {
    await deps.access(helperPath, fsSync.constants.X_OK);
  } catch (error) {
    if (error && typeof error === 'object' && error.code === 'ENOENT') {
      throw new Error(
        `macOS packaging requires ${MACOS_WINDOW_HELPER} at ${helperPath}. Run bun run build:assets on macOS with swiftc available before packaging.`,
      );
    }
    if (error && typeof error === 'object' && error.code === 'EACCES') {
      throw new Error(`macOS window helper is not executable: ${helperPath}`);
    }
    throw error;
  }

  return true;
}

async function afterPack(context) {
  await stageLinuxAppImageSharedLibrary(context);
  await verifyMacOSWindowHelper(context);
}

module.exports = {
  LINUX_FFMPEG_LIBRARY,
  MACOS_WINDOW_HELPER,
  resolveMacOSAppBundlePath,
  stageLinuxAppImageSharedLibrary,
  verifyMacOSWindowHelper,
  default: afterPack,
};
