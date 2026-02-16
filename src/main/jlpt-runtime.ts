import * as path from "path";
import type { JlptLevel } from "../types";

import { createJlptVocabularyLookupService } from "../core/services";

export interface JlptDictionarySearchPathDeps {
  getDictionaryRoots: () => string[];
}

export type JlptLookup = (term: string) => JlptLevel | null;

export interface JlptDictionaryRuntimeDeps {
  isJlptEnabled: () => boolean;
  getSearchPaths: () => string[];
  setJlptLevelLookup: (lookup: JlptLookup) => void;
  log: (message: string) => void;
}

let jlptDictionaryLookupInitialized = false;
let jlptDictionaryLookupInitialization: Promise<void> | null = null;

export function getJlptDictionarySearchPaths(
  deps: JlptDictionarySearchPathDeps,
): string[] {
  const dictionaryRoots = deps.getDictionaryRoots();

  const searchPaths: string[] = [];
  for (const dictionaryRoot of dictionaryRoots) {
    searchPaths.push(dictionaryRoot);
    searchPaths.push(path.join(dictionaryRoot, "vendor", "yomitan-jlpt-vocab"));
    searchPaths.push(path.join(dictionaryRoot, "yomitan-jlpt-vocab"));
  }

  const uniquePaths = new Set<string>(searchPaths);
  return [...uniquePaths];
}

export async function initializeJlptDictionaryLookup(
  deps: JlptDictionaryRuntimeDeps,
): Promise<void> {
  deps.setJlptLevelLookup(
    await createJlptVocabularyLookupService({
      searchPaths: deps.getSearchPaths(),
      log: deps.log,
    }),
  );
}

export async function ensureJlptDictionaryLookup(
  deps: JlptDictionaryRuntimeDeps,
): Promise<void> {
  if (!deps.isJlptEnabled()) {
    return;
  }
  if (jlptDictionaryLookupInitialized) {
    return;
  }
  if (!jlptDictionaryLookupInitialization) {
    jlptDictionaryLookupInitialization = initializeJlptDictionaryLookup(deps)
      .then(() => {
        jlptDictionaryLookupInitialized = true;
      })
      .catch((error) => {
        jlptDictionaryLookupInitialization = null;
        throw error;
      });
  }
  await jlptDictionaryLookupInitialization;
}

export function createJlptDictionaryRuntimeService(
  deps: JlptDictionaryRuntimeDeps,
): { ensureJlptDictionaryLookup: () => Promise<void> } {
  return {
    ensureJlptDictionaryLookup: () => ensureJlptDictionaryLookup(deps),
  };
}
