import fs from 'node:fs';
import path from 'node:path';
import { writeTextFileAtomicallyDurable } from '../shared/fs-utils';

export const SUPPORTED_ELECTRON_MAJOR = 43;
const RUNTIME_STATE_FILE_NAME = 'electron-runtime.json';

type ElectronRuntimeState = {
  highestElectronMajor: number;
  lastElectronVersion: string;
};

type ParsedElectronVersion = {
  major: number;
  minor: number;
  patch: number;
};

type ValidatedElectronRuntimeState = {
  state: ElectronRuntimeState;
  version: ParsedElectronVersion;
};

export type ElectronRuntimeGuardResult =
  | { ok: true; statePath: string }
  | { ok: false; title: string; details: string; statePath: string };

function parseElectronVersion(version: string): ParsedElectronVersion | null {
  const match = /^(\d+)\.(\d+)\.(\d+)(?:\+[0-9A-Za-z.-]+)?$/.exec(version.trim());
  if (!match) return null;

  const major = Number.parseInt(match[1]!, 10);
  const minor = Number.parseInt(match[2]!, 10);
  const patch = Number.parseInt(match[3]!, 10);
  if (![major, minor, patch].every((part) => Number.isSafeInteger(part) && part >= 0)) {
    return null;
  }
  if (major === 0) return null;
  return { major, minor, patch };
}

function compareElectronVersions(
  left: ParsedElectronVersion,
  right: ParsedElectronVersion,
): number {
  return left.major - right.major || left.minor - right.minor || left.patch - right.patch;
}

function readRuntimeState(statePath: string): ValidatedElectronRuntimeState | null {
  if (!fs.existsSync(statePath)) return null;

  const parsed = JSON.parse(fs.readFileSync(statePath, 'utf8')) as Partial<ElectronRuntimeState>;
  const highestElectronMajor = parsed.highestElectronMajor;
  const lastElectronVersion =
    typeof parsed.lastElectronVersion === 'string'
      ? parseElectronVersion(parsed.lastElectronVersion)
      : null;
  if (
    typeof highestElectronMajor !== 'number' ||
    !Number.isSafeInteger(highestElectronMajor) ||
    highestElectronMajor <= 0 ||
    lastElectronVersion === null
  ) {
    throw new Error('The runtime safety record has an invalid format.');
  }

  return {
    state: {
      highestElectronMajor,
      lastElectronVersion: parsed.lastElectronVersion!,
    },
    version: lastElectronVersion,
  };
}

function writeRuntimeState(statePath: string, state: ElectronRuntimeState): void {
  writeTextFileAtomicallyDurable(statePath, `${JSON.stringify(state, null, 2)}\n`);
}

export function enforceElectronRuntimeGuard(options: {
  electronVersion: string;
  userDataPath: string;
  supportedElectronMajor?: number;
}): ElectronRuntimeGuardResult {
  const supportedElectronMajor = options.supportedElectronMajor ?? SUPPORTED_ELECTRON_MAJOR;
  const statePath = path.join(options.userDataPath, RUNTIME_STATE_FILE_NAME);
  const currentVersion = parseElectronVersion(options.electronVersion);

  if (currentVersion === null) {
    return {
      ok: false,
      title: 'SubMiner could not verify Electron',
      details: `Electron reported an invalid version: ${JSON.stringify(options.electronVersion)}. SubMiner did not load Yomitan storage.`,
      statePath,
    };
  }

  if (currentVersion.major !== supportedElectronMajor) {
    return {
      ok: false,
      title: 'Unsupported Electron runtime',
      details: [
        `This SubMiner build requires Electron ${supportedElectronMajor}.`,
        `The current runtime is Electron ${options.electronVersion}.`,
        '',
        'Launch SubMiner through its packaged application or the repository package scripts. Yomitan storage was not loaded.',
      ].join('\n'),
      statePath,
    };
  }

  let previousState: ValidatedElectronRuntimeState | null;
  try {
    previousState = readRuntimeState(statePath);
  } catch (error) {
    return {
      ok: false,
      title: 'SubMiner profile safety check failed',
      details: [
        `SubMiner could not read the runtime safety record at ${statePath}.`,
        (error as Error).message,
        '',
        'Yomitan storage was not loaded. Repair or remove only this safety record after verifying the profile backup.',
      ].join('\n'),
      statePath,
    };
  }

  if (
    previousState &&
    (currentVersion.major < previousState.state.highestElectronMajor ||
      compareElectronVersions(currentVersion, previousState.version) < 0)
  ) {
    return {
      ok: false,
      title: 'Electron downgrade blocked',
      details: [
        `This profile was previously opened with Electron ${previousState.state.lastElectronVersion}.`,
        `The current runtime is Electron ${options.electronVersion}.`,
        '',
        'Opening Chromium storage with an older Electron version can destroy Yomitan dictionaries. Upgrade SubMiner before using this profile.',
      ].join('\n'),
      statePath,
    };
  }

  try {
    writeRuntimeState(statePath, {
      highestElectronMajor: Math.max(
        currentVersion.major,
        previousState?.state.highestElectronMajor ?? 0,
      ),
      lastElectronVersion: options.electronVersion,
    });
  } catch (error) {
    return {
      ok: false,
      title: 'SubMiner profile safety check failed',
      details: [
        `SubMiner could not update the runtime safety record at ${statePath}.`,
        (error as Error).message,
        '',
        'Yomitan storage was not loaded.',
      ].join('\n'),
      statePath,
    };
  }

  return { ok: true, statePath };
}
