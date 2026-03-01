import { MergedToken } from '../../../types';

type MecabPosMetadata = {
  pos1: string;
  pos2?: string;
  pos3?: string;
};

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

function pickClosestMecabPosMetadata(
  token: MergedToken,
  mecabTokens: MergedToken[],
): MecabPosMetadata | null {
  if (mecabTokens.length === 0) {
    return null;
  }

  const tokenStart = token.startPos ?? 0;
  const tokenEnd = token.endPos ?? tokenStart + token.surface.length;
  let bestSurfaceMatchToken: MergedToken | null = null;
  let bestSurfaceMatchDistance = Number.MAX_SAFE_INTEGER;
  let bestSurfaceMatchEndDistance = Number.MAX_SAFE_INTEGER;

  for (const mecabToken of mecabTokens) {
    if (!mecabToken.pos1) {
      continue;
    }

    if (mecabToken.surface !== token.surface) {
      continue;
    }

    const mecabStart = mecabToken.startPos ?? 0;
    const mecabEnd = mecabToken.endPos ?? mecabStart + mecabToken.surface.length;
    const startDistance = Math.abs(mecabStart - tokenStart);
    const endDistance = Math.abs(mecabEnd - tokenEnd);

    if (
      startDistance < bestSurfaceMatchDistance ||
      (startDistance === bestSurfaceMatchDistance && endDistance < bestSurfaceMatchEndDistance)
    ) {
      bestSurfaceMatchDistance = startDistance;
      bestSurfaceMatchEndDistance = endDistance;
      bestSurfaceMatchToken = mecabToken;
    }
  }

  if (bestSurfaceMatchToken) {
    return {
      pos1: bestSurfaceMatchToken.pos1 as string,
      pos2: bestSurfaceMatchToken.pos2,
      pos3: bestSurfaceMatchToken.pos3,
    };
  }

  let bestToken: MergedToken | null = null;
  let bestOverlap = 0;
  let bestSpan = 0;
  let bestStartDistance = Number.MAX_SAFE_INTEGER;
  let bestStart = Number.MAX_SAFE_INTEGER;
  const overlappingTokens: MergedToken[] = [];

  for (const mecabToken of mecabTokens) {
    if (!mecabToken.pos1) {
      continue;
    }

    const mecabStart = mecabToken.startPos ?? 0;
    const mecabEnd = mecabToken.endPos ?? mecabStart + mecabToken.surface.length;
    const overlapStart = Math.max(tokenStart, mecabStart);
    const overlapEnd = Math.min(tokenEnd, mecabEnd);
    const overlap = Math.max(0, overlapEnd - overlapStart);
    if (overlap === 0) {
      continue;
    }
    overlappingTokens.push(mecabToken);

    const span = mecabEnd - mecabStart;
    if (
      overlap > bestOverlap ||
      (overlap === bestOverlap &&
        (Math.abs(mecabStart - tokenStart) < bestStartDistance ||
          (Math.abs(mecabStart - tokenStart) === bestStartDistance &&
            (span > bestSpan || (span === bestSpan && mecabStart < bestStart)))))
    ) {
      bestOverlap = overlap;
      bestSpan = span;
      bestStartDistance = Math.abs(mecabStart - tokenStart);
      bestStart = mecabStart;
      bestToken = mecabToken;
    }
  }

  if (bestOverlap === 0 || !bestToken) {
    return null;
  }

  const overlapPos1 = joinUniqueTags(overlappingTokens.map((token) => token.pos1));
  const overlapPos2 = joinUniqueTags(overlappingTokens.map((token) => token.pos2));
  const overlapPos3 = joinUniqueTags(overlappingTokens.map((token) => token.pos3));

  return {
    pos1: overlapPos1 ?? (bestToken.pos1 as string),
    pos2: overlapPos2 ?? bestToken.pos2,
    pos3: overlapPos3 ?? bestToken.pos3,
  };
}

function fillMissingPos1BySurfaceSequence(
  tokens: MergedToken[],
  mecabTokens: MergedToken[],
): MergedToken[] {
  const indexedMecabTokens = mecabTokens
    .map((token, index) => ({ token, index }))
    .filter(({ token }) => token.pos1 && token.surface.trim().length > 0);

  if (indexedMecabTokens.length === 0) {
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

    let best: { token: MergedToken; index: number } | null = null;
    for (const candidate of indexedMecabTokens) {
      if (candidate.token.surface !== surface) {
        continue;
      }
      if (candidate.index < cursor) {
        continue;
      }
      best = { token: candidate.token, index: candidate.index };
      break;
    }

    if (!best) {
      for (const candidate of indexedMecabTokens) {
        if (candidate.token.surface !== surface) {
          continue;
        }
        best = { token: candidate.token, index: candidate.index };
        break;
      }
    }

    if (!best) {
      return token;
    }

    cursor = best.index + 1;
    return {
      ...token,
      pos1: best.token.pos1,
      pos2: best.token.pos2,
      pos3: best.token.pos3,
    };
  });
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

  const overlapEnriched = tokens.map((token) => {
    if (token.pos1) {
      return token;
    }

    const metadata = pickClosestMecabPosMetadata(token, mecabTokens);
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

  return fillMissingPos1BySurfaceSequence(overlapEnriched, mecabTokens);
}
