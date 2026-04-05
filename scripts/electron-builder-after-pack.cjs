const fs = require('node:fs/promises');
const path = require('node:path');

const LINUX_FFMPEG_LIBRARY = 'libffmpeg.so';

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
  } catch {
    return false;
  }

  const targetLibraryDir = path.join(context.appOutDir, 'usr', 'lib');
  const targetLibraryPath = path.join(targetLibraryDir, LINUX_FFMPEG_LIBRARY);

  await deps.mkdir(targetLibraryDir);
  await deps.copyFile(sourceLibraryPath, targetLibraryPath);

  return true;
}

async function afterPack(context) {
  await stageLinuxAppImageSharedLibrary(context);
}

module.exports = {
  LINUX_FFMPEG_LIBRARY,
  stageLinuxAppImageSharedLibrary,
  default: afterPack,
};
