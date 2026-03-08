import electron from 'electron';
import type { BrowserWindow, Extension } from 'electron';
import * as fs from 'fs';
import { createLogger } from '../../logger';
import { ensureExtensionCopy } from './yomitan-extension-copy';
import {
  getYomitanExtensionSearchPaths,
  resolveExistingYomitanExtensionPath,
} from './yomitan-extension-paths';

const { session } = electron;
const logger = createLogger('main:yomitan-extension-loader');

export interface YomitanExtensionLoaderDeps {
  userDataPath: string;
  extensionPath?: string;
  getYomitanParserWindow: () => BrowserWindow | null;
  setYomitanParserWindow: (window: BrowserWindow | null) => void;
  setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
  setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
  setYomitanExtension: (extension: Extension | null) => void;
}

export async function loadYomitanExtension(
  deps: YomitanExtensionLoaderDeps,
): Promise<Extension | null> {
  const searchPaths = getYomitanExtensionSearchPaths({
    explicitPath: deps.extensionPath,
    moduleDir: __dirname,
    resourcesPath: process.resourcesPath,
    userDataPath: deps.userDataPath,
  });
  let extPath = resolveExistingYomitanExtensionPath(searchPaths, fs.existsSync);

  if (!extPath) {
    logger.error('Yomitan extension not found in any search path');
    logger.error('Run `bun run build:yomitan` or install Yomitan to one of:', searchPaths);
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
