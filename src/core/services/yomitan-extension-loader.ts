import { BrowserWindow, Extension, session } from 'electron';
import * as fs from 'fs';
import * as path from 'path';
import { createLogger } from '../../logger';
import { ensureExtensionCopy } from './yomitan-extension-copy';

const logger = createLogger('main:yomitan-extension-loader');

export interface YomitanExtensionLoaderDeps {
  userDataPath: string;
  getYomitanParserWindow: () => BrowserWindow | null;
  setYomitanParserWindow: (window: BrowserWindow | null) => void;
  setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
  setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
  setYomitanExtension: (extension: Extension | null) => void;
}

export async function loadYomitanExtension(
  deps: YomitanExtensionLoaderDeps,
): Promise<Extension | null> {
  const searchPaths = [
    path.join(__dirname, '..', '..', 'vendor', 'yomitan'),
    path.join(__dirname, '..', '..', '..', 'vendor', 'yomitan'),
    path.join(process.resourcesPath, 'yomitan'),
    '/usr/share/SubMiner/yomitan',
    path.join(deps.userDataPath, 'yomitan'),
  ];

  let extPath: string | null = null;
  for (const p of searchPaths) {
    if (fs.existsSync(p)) {
      extPath = p;
      break;
    }
  }

  if (!extPath) {
    logger.error('Yomitan extension not found in any search path');
    logger.error('Install Yomitan to one of:', searchPaths);
    return null;
  }

  const extensionCopy = ensureExtensionCopy(extPath, deps.userDataPath);
  if (extensionCopy.copied) {
    logger.info(`Copied yomitan extension to ${extensionCopy.targetDir}`);
  }
  extPath = extensionCopy.targetDir;

  const parserWindow = deps.getYomitanParserWindow();
  if (parserWindow && !parserWindow.isDestroyed()) {
    parserWindow.destroy();
  }
  deps.setYomitanParserWindow(null);
  deps.setYomitanParserReadyPromise(null);
  deps.setYomitanParserInitPromise(null);

  try {
    const extensions = session.defaultSession.extensions;
    const extension = extensions
      ? await extensions.loadExtension(extPath, {
          allowFileAccess: true,
        })
      : await session.defaultSession.loadExtension(extPath, {
          allowFileAccess: true,
        });
    deps.setYomitanExtension(extension);
    return extension;
  } catch (err) {
    logger.error('Failed to load Yomitan extension:', (err as Error).message);
    logger.error('Full error:', err);
    deps.setYomitanExtension(null);
    return null;
  }
}
