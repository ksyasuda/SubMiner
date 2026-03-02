import { MergedToken } from '../../../types';

type MecabPosMetadata = {
  pos1: string;
  pos2?: string;
  pos3?: string;
};

type IndexedMecabToken = {
  index: number;
  start: number;
  end: number;
  surface: string;
  pos1: string;
  pos2?: string;
  pos3?: string;
};

type MecabLookup = {
  indexedTokens: IndexedMecabToken[];
  byExactSurface: Map<string, IndexedMecabToken[]>;
  byTrimmedSurface: Map<string, IndexedMecabToken[]>;
  byPosition: Map<number, IndexedMecabToken[]>;
};

function pushMapValue<K, T>(map: Map<K, T[]>, key: K, value: T): void {
  const existing = map.get(key);
  if (existing) {
    existing.push(value);
    return;
  }
  map.set(key, [value]);
}

function toDiscreteSpan(start: number, end: number): { start: number; end: number } {
  const discreteStart = Math.floor(start);
  const discreteEnd = Math.max(discreteStart + 1, Math.ceil(end));
  return {
    start: discreteStart,
    end: discreteEnd,
  };
}

function buildMecabLookup(mecabTokens: MergedToken[]): MecabLookup {
  const indexedTokens: IndexedMecabToken[] = [];
  for (const [index, token] of mecabTokens.entries()) {
    const pos1 = token.pos1;
    if (!pos1) {
      continue;
    }
    const surface = token.surface;
    const start = token.startPos ?? 0;
    const end = token.endPos ?? start + surface.length;
    indexedTokens.push({
      index,
      start,
      end,
      surface,
      pos1,
      pos2: token.pos2,
      pos3: token.pos3,
    });
  }

  const byExactSurface = new Map<string, IndexedMecabToken[]>();
  const byTrimmedSurface = new Map<string, IndexedMecabToken[]>();
  const byPosition = new Map<number, IndexedMecabToken[]>();
  for (const token of indexedTokens) {
    pushMapValue(byExactSurface, token.surface, token);
    const trimmedSurface = token.surface.trim();
    if (trimmedSurface) {
      pushMapValue(byTrimmedSurface, trimmedSurface, token);
    }

    const discreteSpan = toDiscreteSpan(token.start, token.end);
    for (let position = discreteSpan.start; position < discreteSpan.end; position += 1) {
      pushMapValue(byPosition, position, token);
    }
  }

  const byStartThenIndexSort = (left: IndexedMecabToken, right: IndexedMecabToken) =>
    left.start - right.start || left.index - right.index;
  for (const candidates of byExactSurface.values()) {
    candidates.sort(byStartThenIndexSort);
  }

  return {
    indexedTokens,
    byExactSurface,
    byTrimmedSurface,
    byPosition,
  };
}

function lowerBoundByStart(candidates: IndexedMecabToken[], targetStart: number): number {
  let low = 0;
  let high = candidates.length;
  while (low < high) {
    const mid = Math.floor((low + high) / 2);
    if (candidates[mid]!.start < targetStart) {
      low = mid + 1;
    } else {
      high = mid;
    }
  }
  return low;
}

function lowerBoundByIndex(candidates: IndexedMecabToken[], targetIndex: number): number {
  let low = 0;
  let high = candidates.length;
  while (low < high) {
    const mid = Math.floor((low + high) / 2);
    if (candidates[mid]!.index < targetIndex) {
      low = mid + 1;
    } else {
      high = mid;
    }
  }
  return low;
}

function joinUniqueTags(values: Array<string | undefined>): string | undefined {
  const unique: string[] = [];
  for (const value of values) {
    if (!value) {
      continue;
    }
    const trimmed = value.trim();
    if (!trimmed) {
      continue;
    }
    if (!unique.includes(trimmed)) {
      unique.push(trimmed);
    }
  }
  if (unique.length === 0) {
    return undefined;
  }
  if (unique.length === 1) {
    return unique[0];
  }
  return unique.join('|');
}

function pickClosestMecabPosMetadataBySurface(
  token: MergedToken,
  candidates: IndexedMecabToken[] | undefined,
): MecabPosMetadata | null {
  if (!candidates || candidates.length === 0) {
    return null;
  }

  const tokenStart = token.startPos ?? 0;
  const tokenEnd = token.endPos ?? tokenStart + token.surface.length;
  let bestSurfaceMatchToken: IndexedMecabToken | null = null;
  let bestSurfaceMatchDistance = Number.MAX_SAFE_INTEGER;
  let bestSurfaceMatchEndDistance = Number.MAX_SAFE_INTEGER;
  let bestSurfaceMatchIndex = Number.MAX_SAFE_INTEGER;

  const nearestStartIndex = lowerBoundByStart(candidates, tokenStart);
  let left = nearestStartIndex - 1;
  let right = nearestStartIndex;

  while (left >= 0 || right < candidates.length) {
    const leftDistance =
      left >= 0 ? Math.abs(candidates[left]!.start - tokenStart) : Number.MAX_SAFE_INTEGER;
    const rightDistance =
      right < candidates.length
        ? Math.abs(candidates[right]!.start - tokenStart)
        : Number.MAX_SAFE_INTEGER;
    const nearestDistance = Math.min(leftDistance, rightDistance);
    if (nearestDistance > bestSurfaceMatchDistance) {
      break;
    }

    if (leftDistance === nearestDistance && left >= 0) {
      const candidate = candidates[left]!;
      const startDistance = Math.abs(candidate.start - tokenStart);
      const endDistance = Math.abs(candidate.end - tokenEnd);
      if (
        startDistance < bestSurfaceMatchDistance ||
        (startDistance === bestSurfaceMatchDistance &&
          (endDistance < bestSurfaceMatchEndDistance ||
            (endDistance === bestSurfaceMatchEndDistance && candidate.index < bestSurfaceMatchIndex)))
      ) {
        bestSurfaceMatchDistance = startDistance;
        bestSurfaceMatchEndDistance = endDistance;
        bestSurfaceMatchIndex = candidate.index;
        bestSurfaceMatchToken = candidate;
      }
      left -= 1;
    }
    if (rightDistance === nearestDistance && right < candidates.length) {
      const candidate = candidates[right]!;
      const startDistance = Math.abs(candidate.start - tokenStart);
      const endDistance = Math.abs(candidate.end - tokenEnd);
      if (
        startDistance < bestSurfaceMatchDistance ||
        (startDistance === bestSurfaceMatchDistance &&
          (endDistance < bestSurfaceMatchEndDistance ||
            (endDistance === bestSurfaceMatchEndDistance && candidate.index < bestSurfaceMatchIndex)))
      ) {
        bestSurfaceMatchDistance = startDistance;
        bestSurfaceMatchEndDistance = endDistance;
        bestSurfaceMatchIndex = candidate.index;
        bestSurfaceMatchToken = candidate;
      }
      right += 1;
    }
  }

  if (bestSurfaceMatchToken !== null) {
    return {
      pos1: bestSurfaceMatchToken.pos1,
      pos2: bestSurfaceMatchToken.pos2,
      pos3: bestSurfaceMatchToken.pos3,
    };
  }

  return null;
}

function pickClosestMecabPosMetadataByOverlap(
  token: MergedToken,
  candidates: IndexedMecabToken[],
): MecabPosMetadata | null {
  const tokenStart = token.startPos ?? 0;
  const tokenEnd = token.endPos ?? tokenStart + token.surface.length;
  let bestToken: IndexedMecabToken | null = null;
  let bestOverlap = 0;
  let bestSpan = 0;
  let bestStartDistance = Number.MAX_SAFE_INTEGER;
  let bestStart = Number.MAX_SAFE_INTEGER;
  let bestIndex = Number.MAX_SAFE_INTEGER;
  const overlappingTokens: IndexedMecabToken[] = [];

  for (const candidate of candidates) {
    const mecabStart = candidate.start;
    const mecabEnd = candidate.end;
    const overlapStart = Math.max(tokenStart, mecabStart);
    const overlapEnd = Math.min(tokenEnd, mecabEnd);
    const overlap = Math.max(0, overlapEnd - overlapStart);
    if (overlap === 0) {
      continue;
    }
    overlappingTokens.push(candidate);

    const span = mecabEnd - mecabStart;
    const startDistance = Math.abs(mecabStart - tokenStart);
    if (
      overlap > bestOverlap ||
      (overlap === bestOverlap &&
        (startDistance < bestStartDistance ||
          (startDistance === bestStartDistance &&
            (span > bestSpan ||
              (span === bestSpan &&
                (mecabStart < bestStart ||
                  (mecabStart === bestStart && candidate.index < bestIndex)))))))
    ) {
      bestOverlap = overlap;
      bestSpan = span;
      bestStartDistance = startDistance;
      bestStart = mecabStart;
      bestIndex = candidate.index;
      bestToken = candidate;
    }
  }

  if (bestOverlap === 0 || !bestToken) {
    return null;
  }

  const overlappingTokensByMecabOrder = overlappingTokens
    .slice()
    .sort((left, right) => left.index - right.index);
  const overlapPos1 = joinUniqueTags(overlappingTokensByMecabOrder.map((candidate) => candidate.pos1));
  const overlapPos2 = joinUniqueTags(overlappingTokensByMecabOrder.map((candidate) => candidate.pos2));
  const overlapPos3 = joinUniqueTags(overlappingTokensByMecabOrder.map((candidate) => candidate.pos3));

  return {
    pos1: overlapPos1 ?? bestToken.pos1,
    pos2: overlapPos2 ?? bestToken.pos2,
    pos3: overlapPos3 ?? bestToken.pos3,
  };
}

function fillMissingPos1BySurfaceSequence(
  tokens: MergedToken[],
  byTrimmedSurface: Map<string, IndexedMecabToken[]>,
): MergedToken[] {
  if (byTrimmedSurface.size === 0) {
    return tokens;
  }

  let cursor = 0;
  return tokens.map((token) => {
    if (token.pos1 && token.pos1.trim().length > 0) {
      return token;
    }

    const surface = token.surface.trim();
    if (!surface) {
      return token;
    }

    const candidates = byTrimmedSurface.get(surface);
    if (!candidates || candidates.length === 0) {
      return token;
    }

    const atOrAfterCursorIndex = lowerBoundByIndex(candidates, cursor);
    const best = candidates[atOrAfterCursorIndex] ?? candidates[0];

    if (!best) {
      return token;
    }

    cursor = best.index + 1;
    return {
      ...token,
      pos1: best.pos1,
      pos2: best.pos2,
      pos3: best.pos3,
    };
  });
}

function collectOverlapCandidatesByPosition(
  token: MergedToken,
  byPosition: Map<number, IndexedMecabToken[]>,
): IndexedMecabToken[] {
  const tokenStart = token.startPos ?? 0;
  const tokenEnd = token.endPos ?? tokenStart + token.surface.length;
  const discreteSpan = toDiscreteSpan(tokenStart, tokenEnd);
  const seen = new Set<number>();
  const overlapCandidates: IndexedMecabToken[] = [];

  for (let position = discreteSpan.start; position < discreteSpan.end; position += 1) {
    const candidatesAtPosition = byPosition.get(position);
    if (!candidatesAtPosition) {
      continue;
    }

    for (const candidate of candidatesAtPosition) {
      if (seen.has(candidate.index)) {
        continue;
      }
      seen.add(candidate.index);
      overlapCandidates.push(candidate);
    }
  }

  return overlapCandidates;
}

export function enrichTokensWithMecabPos1(
  tokens: MergedToken[],
  mecabTokens: MergedToken[] | null,
): MergedToken[] {
  if (!tokens || tokens.length === 0) {
    return tokens;
  }

  if (!mecabTokens || mecabTokens.length === 0) {
    return tokens;
  }

  const lookup = buildMecabLookup(mecabTokens);
  if (lookup.indexedTokens.length === 0) {
    return tokens;
  }

  const metadataByTokenIndex = new Map<number, MecabPosMetadata>();

  for (const [index, token] of tokens.entries()) {
    if (token.pos1) {
      continue;
    }

    const surfaceMetadata = pickClosestMecabPosMetadataBySurface(
      token,
      lookup.byExactSurface.get(token.surface),
    );
    if (surfaceMetadata) {
      metadataByTokenIndex.set(index, surfaceMetadata);
      continue;
    }

    const overlapCandidates = collectOverlapCandidatesByPosition(token, lookup.byPosition);
    const overlapMetadata = pickClosestMecabPosMetadataByOverlap(token, overlapCandidates);
    if (overlapMetadata) {
      metadataByTokenIndex.set(index, overlapMetadata);
    }
  }

  const overlapEnriched = tokens.map((token, index) => {
    const metadata = metadataByTokenIndex.get(index);
    if (!metadata) {
      return token;
    }

    return {
      ...token,
      pos1: metadata.pos1,
      pos2: metadata.pos2,
      pos3: metadata.pos3,
    };
  });

  return fillMissingPos1BySurfaceSequence(overlapEnriched, lookup.byTrimmedSurface);
}
