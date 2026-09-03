import fs from 'node:fs';
import path from 'node:path';
import { writeTextFileAtomicallyDurable } from '../../shared/fs-utils';
import { getSetupStatePath, readSetupState } from '../../shared/setup-state';

const INTEGRITY_STATE_FILE_NAME = 'yomitan-dictionary-integrity.json';

type DictionaryIntegrityState = {
  lastKnownNonEmptyCount: number;
};

export type DictionaryIntegrityObservation =
  | { safe: true; previousCount: number | null }
  | { safe: false; previousCount: number | null; message: string };

function getStatePath(userDataPath: string): string {
  return path.join(userDataPath, INTEGRITY_STATE_FILE_NAME);
}

function readState(statePath: string): DictionaryIntegrityState | null {
  if (!fs.existsSync(statePath)) return null;
  const parsed = JSON.parse(
    fs.readFileSync(statePath, 'utf8'),
  ) as Partial<DictionaryIntegrityState>;
  const lastKnownNonEmptyCount = parsed.lastKnownNonEmptyCount;
  if (
    typeof lastKnownNonEmptyCount !== 'number' ||
    !Number.isSafeInteger(lastKnownNonEmptyCount) ||
    lastKnownNonEmptyCount <= 0
  ) {
    throw new Error('The dictionary integrity record has an invalid format.');
  }
  return { lastKnownNonEmptyCount };
}

function writeState(statePath: string, state: DictionaryIntegrityState): void {
  writeTextFileAtomicallyDurable(statePath, `${JSON.stringify(state, null, 2)}\n`);
}

function readLegacySetupCount(userDataPath: string): number | null {
  const setupState = readSetupState(getSetupStatePath(userDataPath));
  return setupState && setupState.lastSeenYomitanDictionaryCount > 0
    ? setupState.lastSeenYomitanDictionaryCount
    : null;
}

export function observeYomitanDictionaryCount(
  userDataPath: string,
  dictionaryCount: number,
): DictionaryIntegrityObservation {
  if (!Number.isSafeInteger(dictionaryCount) || dictionaryCount < 0) {
    return {
      safe: false,
      previousCount: null,
      message: 'SubMiner could not verify Yomitan dictionary storage: invalid dictionary count.',
    };
  }

  const normalizedCount = dictionaryCount;
  const statePath = getStatePath(userDataPath);
  let state: DictionaryIntegrityState | null;

  try {
    state = readState(statePath);
    if (state === null) {
      const legacyCount = readLegacySetupCount(userDataPath);
      state = legacyCount === null ? null : { lastKnownNonEmptyCount: legacyCount };
    }
  } catch (error) {
    return {
      safe: false,
      previousCount: null,
      message: `SubMiner could not verify Yomitan dictionary storage at ${statePath}: ${(error as Error).message}`,
    };
  }

  if (normalizedCount === 0 && state !== null) {
    return {
      safe: false,
      previousCount: state.lastKnownNonEmptyCount,
      message: [
        `Yomitan reported zero dictionaries after previously reporting ${state.lastKnownNonEmptyCount}.`,
        'SubMiner blocked automatic dictionary changes because Chromium storage may have been reset.',
        'Close SubMiner and restore or inspect the profile before changing dictionaries.',
      ].join(' '),
    };
  }

  if (normalizedCount > 0 && normalizedCount !== state?.lastKnownNonEmptyCount) {
    try {
      writeState(statePath, { lastKnownNonEmptyCount: normalizedCount });
    } catch (error) {
      return {
        safe: false,
        previousCount: state?.lastKnownNonEmptyCount ?? null,
        message: `SubMiner could not update the Yomitan dictionary integrity record: ${(error as Error).message}`,
      };
    }
  }

  return { safe: true, previousCount: state?.lastKnownNonEmptyCount ?? null };
}

export function assertYomitanDictionaryMutationSafe(
  userDataPath: string,
  dictionaryCount: number,
): void {
  const observation = observeYomitanDictionaryCount(userDataPath, dictionaryCount);
  if (!observation.safe) {
    throw new Error(observation.message);
  }
}
