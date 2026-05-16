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

type WarningProcess = Pick<NodeJS.Process, 'emitWarning'>;

function getWarningType(warning: string | Error, args: unknown[]): string | undefined {
  if (typeof warning !== 'string') {
    return warning.name;
  }
  const firstArg = args[0];
  if (typeof firstArg === 'string') {
    return firstArg;
  }
  if (firstArg && typeof firstArg === 'object' && 'type' in firstArg) {
    const type = (firstArg as { type?: unknown }).type;
    return typeof type === 'string' ? type : undefined;
  }
  return undefined;
}

function shouldSuppressYomitanExtensionWarning(warning: string | Error, args: unknown[]): boolean {
  const message = warning instanceof Error ? warning.message : warning;
  return (
    getWarningType(warning, args) === 'ExtensionLoadWarning' &&
    message.includes("Permission 'contextMenus' is unknown.")
  );
}

export async function withSuppressedYomitanExtensionWarnings<T>(
  run: () => Promise<T>,
  warningProcess: WarningProcess = process,
): Promise<T> {
  const originalEmitWarning = warningProcess.emitWarning;
  warningProcess.emitWarning = ((warning: string | Error, ...args: unknown[]) => {
    if (shouldSuppressYomitanExtensionWarning(warning, args)) {
      return;
    }
    return (originalEmitWarning as (...emitArgs: unknown[]) => void).call(
      warningProcess,
      warning,
      ...args,
    );
  }) as typeof process.emitWarning;

  try {
    return await run();
  } finally {
    warningProcess.emitWarning = originalEmitWarning;
  }
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
      logger.debug(`Copied yomitan extension to ${extensionCopy.targetDir}`);
    }
    extPath = extensionCopy.targetDir;
  }

  clearParserState();
  deps.setYomitanSession(targetSession);

  try {
    const extensions = targetSession.extensions;
    const extension = await withSuppressedYomitanExtensionWarnings(() =>
      extensions
        ? extensions.loadExtension(extPath, {
            allowFileAccess: true,
          })
        : targetSession.loadExtension(extPath, {
            allowFileAccess: true,
          }),
    );
    deps.setYomitanExtension(extension);
    return extension;
  } catch (err) {
    logger.error('Failed to load Yomitan extension:', (err as Error).message);
    logger.error('Full error:', err);
    clearRuntimeState();
    return null;
  }
}
