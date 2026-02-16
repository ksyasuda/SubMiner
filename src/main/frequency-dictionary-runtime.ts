import * as path from "path";
import type { FrequencyDictionaryLookup } from "../types";
import { createFrequencyDictionaryLookupService } from "../core/services";

export interface FrequencyDictionarySearchPathDeps {
  getDictionaryRoots: () => string[];
  getSourcePath?: () => string | undefined;
}

export interface FrequencyDictionaryRuntimeDeps {
  isFrequencyDictionaryEnabled: () => boolean;
  getSearchPaths: () => string[];
  setFrequencyRankLookup: (lookup: FrequencyDictionaryLookup) => void;
  log: (message: string) => void;
}

let frequencyDictionaryLookupInitialized = false;
let frequencyDictionaryLookupInitialization: Promise<void> | null = null;

export function getFrequencyDictionarySearchPaths(
  deps: FrequencyDictionarySearchPathDeps,
): string[] {
  const dictionaryRoots = deps.getDictionaryRoots();
  const sourcePath = deps.getSourcePath?.();

  const rawSearchPaths: string[] = [];
  if (sourcePath && sourcePath.trim()) {
    rawSearchPaths.push(sourcePath.trim());
    rawSearchPaths.push(path.join(sourcePath.trim(), "frequency-dictionary"));
    rawSearchPaths.push(path.join(sourcePath.trim(), "vendor", "frequency-dictionary"));
  }

  for (const dictionaryRoot of dictionaryRoots) {
    rawSearchPaths.push(dictionaryRoot);
    rawSearchPaths.push(path.join(dictionaryRoot, "frequency-dictionary"));
    rawSearchPaths.push(path.join(dictionaryRoot, "vendor", "frequency-dictionary"));
  }

  return [...new Set(rawSearchPaths)];
}

export async function initializeFrequencyDictionaryLookup(
  deps: FrequencyDictionaryRuntimeDeps,
): Promise<void> {
  const lookup = await createFrequencyDictionaryLookupService({
    searchPaths: deps.getSearchPaths(),
    log: deps.log,
  });
  deps.setFrequencyRankLookup(lookup);
}

export async function ensureFrequencyDictionaryLookup(
  deps: FrequencyDictionaryRuntimeDeps,
): Promise<void> {
  if (!deps.isFrequencyDictionaryEnabled()) {
    return;
  }
  if (frequencyDictionaryLookupInitialized) {
    return;
  }
  if (!frequencyDictionaryLookupInitialization) {
    frequencyDictionaryLookupInitialization = initializeFrequencyDictionaryLookup(deps)
      .then(() => {
        frequencyDictionaryLookupInitialized = true;
      })
      .catch((error) => {
        frequencyDictionaryLookupInitialization = null;
        throw error;
      });
  }
  await frequencyDictionaryLookupInitialization;
}

export function createFrequencyDictionaryRuntimeService(
  deps: FrequencyDictionaryRuntimeDeps,
): { ensureFrequencyDictionaryLookup: () => Promise<void> } {
  return {
    ensureFrequencyDictionaryLookup: () => ensureFrequencyDictionaryLookup(deps),
  };
}
