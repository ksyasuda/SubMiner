import electron from 'electron';
import type { BrowserWindow, Extension, Session } from 'electron';
import * as fs from 'fs';
import * as path from 'path';
import { createLogger } from '../../logger';
import { ensureExtensionCopyAsync } from './yomitan-extension-copy';
import {
  getYomitanExtensionSearchPaths,
  resolveExternalYomitanExtensionPath,
  resolveExistingYomitanExtensionPath,
} from './yomitan-extension-paths';
import {
  clearYomitanExtensionRuntimeState,
  clearYomitanParserRuntimeState,
} from './yomitan-extension-runtime-state';

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
  const clearRuntimeState = () =>
    clearYomitanExtensionRuntimeState({
      getYomitanParserWindow: deps.getYomitanParserWindow,
      setYomitanParserWindow: deps.setYomitanParserWindow,
      setYomitanParserReadyPromise: deps.setYomitanParserReadyPromise,
      setYomitanParserInitPromise: deps.setYomitanParserInitPromise,
      setYomitanExtension: () => deps.setYomitanExtension(null),
      setYomitanSession: () => deps.setYomitanSession(null),
    });
  const clearParserState = () =>
    clearYomitanParserRuntimeState({
      getYomitanParserWindow: deps.getYomitanParserWindow,
      setYomitanParserWindow: deps.setYomitanParserWindow,
      setYomitanParserReadyPromise: deps.setYomitanParserReadyPromise,
      setYomitanParserInitPromise: deps.setYomitanParserInitPromise,
    });
  const externalProfilePath = deps.externalProfilePath?.trim() ?? '';
  let extPath: string | null = null;
  let targetSession: Session = session.defaultSession;

  if (externalProfilePath) {
    const resolvedProfilePath = path.resolve(externalProfilePath);
    extPath = resolveExternalYomitanExtensionPath(resolvedProfilePath, fs.existsSync);
    if (!extPath) {
      logger.error('External Yomitan extension not found in configured profile path');
      logger.error('Expected unpacked extension at:', path.join(resolvedProfilePath, 'extensions'));
      clearRuntimeState();
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
      clearRuntimeState();
      return null;
    }

    const extensionCopy = await ensureExtensionCopyAsync(extPath, deps.userDataPath);
    if (extensionCopy.copied) {
      logger.info(`Copied yomitan extension to ${extensionCopy.targetDir}`);
    }
    extPath = extensionCopy.targetDir;
  }

  clearParserState();
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
    clearRuntimeState();
    return null;
  }
}
