import * as fs from 'fs';
import * as path from 'path';
import { readCachedSnapshots } from './cache';
import type { CharacterDictionarySnapshot } from './types';

// Candidate name forms for the greedy name pre-pass in the Yomitan scan
// runtime. The scanner otherwise has to ask the backend at every Japanese
// position, because a character name can start mid-token; knowing which forms
// exist lets it look up only where a name can actually begin.
//
// A form is any string Yomitan could match a character entry by: the term and
// its reading. Both come from the dictionary SubMiner generated, so the pair is
// the complete matchable set for an entry. Callers treat a missing list as
// "scan every position", so a stale or absent snapshot costs speed, never a
// missed name.

function getSnapshotsDir(outputDir: string): string {
  return path.join(outputDir, 'snapshots');
}

function collectSnapshotNameForms(snapshot: CharacterDictionarySnapshot): string[] {
  const forms = new Set<string>();
  for (const entry of snapshot.termEntries) {
    const term = typeof entry[0] === 'string' ? entry[0].trim() : '';
    if (term) {
      forms.add(term);
    }
    const reading = typeof entry[1] === 'string' ? entry[1].trim() : '';
    if (reading) {
      forms.add(reading);
    }
  }
  return [...forms];
}

// The signature grows with the size of the dictionary library, and it rides
// along in every per-line scan call, so it is folded into a fixed-width digest
// first. Collisions only matter against the immediately previous signature (the
// runtime compares keys for equality), and FNV-1a over the file list is far
// beyond what that needs.
function digestSnapshotDirectorySignature(signature: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < signature.length; index += 1) {
    hash ^= signature.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(36);
}

function getSnapshotDirectorySignature(outputDir: string): string {
  let entries: fs.Dirent[] = [];
  try {
    entries = fs.readdirSync(getSnapshotsDir(outputDir), { withFileTypes: true });
  } catch {
    return '';
  }

  const parts: string[] = [];
  for (const entry of entries) {
    if (!entry.isFile() || !/^anilist-\d+\.json$/.test(entry.name)) {
      continue;
    }
    try {
      const stat = fs.statSync(path.join(getSnapshotsDir(outputDir), entry.name));
      parts.push(`${entry.name}:${stat.mtimeMs}:${stat.size}`);
    } catch {
      // Ignore files that disappear during a refresh; the next lookup rebuilds.
    }
  }
  return parts.sort().join('|');
}

export interface CharacterNameCandidateSet {
  /** Identifies this exact form list, so the scan runtime can cache it. */
  key: string;
  forms: string[];
}

// This lookup is consulted once per subtitle line, so it must not stat the
// snapshot directory every time. Dictionary writes are rare and always call
// invalidate(), which forces the next lookup to re-read; the interval only
// bounds staleness from changes made behind our back.
const SNAPSHOT_SIGNATURE_RECHECK_INTERVAL_MS = 5000;

export function createCharacterNameCandidateLookup(deps: {
  userDataPath?: string;
  outputDir?: string;
  getCurrentMediaId?: () => number | null | undefined;
  now?: () => number;
}): {
  get: (mediaId?: number | null) => CharacterNameCandidateSet | null;
  invalidate: () => void;
} {
  const outputDir =
    deps.outputDir ??
    (deps.userDataPath ? path.join(deps.userDataPath, 'character-dictionaries') : '');
  const now = deps.now ?? (() => Date.now());
  let signature: string | null = null;
  let lastSignatureCheckAtMs = 0;
  let formsByMediaId = new Map<number, string[]>();

  function refreshIfNeeded(): void {
    if (!outputDir) {
      formsByMediaId = new Map<number, string[]>();
      signature = '';
      return;
    }
    const nowMs = now();
    if (
      signature !== null &&
      nowMs - lastSignatureCheckAtMs < SNAPSHOT_SIGNATURE_RECHECK_INTERVAL_MS
    ) {
      return;
    }
    lastSignatureCheckAtMs = nowMs;
    const nextSignature = getSnapshotDirectorySignature(outputDir);
    if (nextSignature === signature) {
      return;
    }
    signature = nextSignature;
    formsByMediaId = new Map<number, string[]>();
    for (const snapshot of readCachedSnapshots(outputDir)) {
      const forms = collectSnapshotNameForms(snapshot);
      if (forms.length > 0) {
        formsByMediaId.set(snapshot.mediaId, forms);
      }
    }
  }

  return {
    get(mediaId?: number | null): CharacterNameCandidateSet | null {
      refreshIfNeeded();
      const rawMediaId = mediaId ?? deps.getCurrentMediaId?.() ?? null;
      const normalizedMediaId =
        typeof rawMediaId === 'number' && Number.isFinite(rawMediaId) && rawMediaId > 0
          ? Math.floor(rawMediaId)
          : null;

      // Without a media scope the pre-pass would need every character of every
      // cached title, which is both slow to match and pointless: report no
      // candidates so the scanner keeps its exhaustive behavior.
      if (normalizedMediaId === null) {
        return null;
      }
      const forms = formsByMediaId.get(normalizedMediaId);
      if (!forms || forms.length === 0) {
        return null;
      }
      return {
        key: `${digestSnapshotDirectorySignature(signature ?? '')}:${normalizedMediaId}`,
        forms,
      };
    },
    invalidate(): void {
      signature = null;
      lastSignatureCheckAtMs = 0;
    },
  };
}
