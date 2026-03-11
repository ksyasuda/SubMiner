import electron from 'electron';
import type { BrowserWindow, Extension, Session } from 'electron';
import * as fs from 'fs';
import * as path from 'path';
import { createLogger } from '../../logger';
import { ensureExtensionCopy } from './yomitan-extension-copy';
import {
  getYomitanExtensionSearchPaths,
  resolveExternalYomitanExtensionPath,
  resolveExistingYomitanExtensionPath,
} from './yomitan-extension-paths';

const { session } = electron;
const logger = createLogger('main:yomitan-extension-loader');

export interface YomitanExtensionLoaderDeps {
  userDataPath: string;
  extensionPath?: string;
  externalProfilePath?: string;
  getYomitanParserWindow: () => BrowserWindow | null;
  setYomitanParserWindow: (window: BrowserWindow | null) => void;
  setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
  setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
  setYomitanExtension: (extension: Extension | null) => void;
  setYomitanSession: (session: Session | null) => void;
}

export async function loadYomitanExtension(
  deps: YomitanExtensionLoaderDeps,
): Promise<Extension | null> {
  const externalProfilePath = deps.externalProfilePath?.trim() ?? '';
  let extPath: string | null = null;
  let targetSession: Session = session.defaultSession;

  if (externalProfilePath) {
    const resolvedProfilePath = path.resolve(externalProfilePath);
    extPath = resolveExternalYomitanExtensionPath(resolvedProfilePath, fs.existsSync);
    if (!extPath) {
      logger.error('External Yomitan extension not found in configured profile path');
      logger.error('Expected unpacked extension at:', path.join(resolvedProfilePath, 'extensions'));
      deps.setYomitanExtension(null);
      deps.setYomitanSession(null);
      return null;
    }

    targetSession = session.fromPath(resolvedProfilePath);
  } else {
    const searchPaths = getYomitanExtensionSearchPaths({
      explicitPath: deps.extensionPath,
      moduleDir: __dirname,
      resourcesPath: process.resourcesPath,
      userDataPath: deps.userDataPath,
    });
    extPath = resolveExistingYomitanExtensionPath(searchPaths, fs.existsSync);

    if (!extPath) {
      logger.error('Yomitan extension not found in any search path');
      logger.error('Run `bun run build:yomitan` or install Yomitan to one of:', searchPaths);
      deps.setYomitanExtension(null);
      deps.setYomitanSession(null);
      return null;
    }

    const extensionCopy = ensureExtensionCopy(extPath, deps.userDataPath);
    if (extensionCopy.copied) {
      logger.info(`Copied yomitan extension to ${extensionCopy.targetDir}`);
    }
    extPath = extensionCopy.targetDir;
  }

  const parserWindow = deps.getYomitanParserWindow();
  if (parserWindow && !parserWindow.isDestroyed()) {
    parserWindow.destroy();
  }
  deps.setYomitanParserWindow(null);
  deps.setYomitanParserReadyPromise(null);
  deps.setYomitanParserInitPromise(null);
  deps.setYomitanSession(targetSession);

  try {
    const extensions = targetSession.extensions;
    const extension = extensions
      ? await extensions.loadExtension(extPath, {
          allowFileAccess: true,
        })
      : await targetSession.loadExtension(extPath, {
          allowFileAccess: true,
        });
    deps.setYomitanExtension(extension);
    return extension;
  } catch (err) {
    logger.error('Failed to load Yomitan extension:', (err as Error).message);
    logger.error('Full error:', err);
    deps.setYomitanExtension(null);
    deps.setYomitanSession(null);
    return null;
  }
}
