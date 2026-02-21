import { MergedToken } from '../../../types';

function pickClosestMecabPos1(token: MergedToken, mecabTokens: MergedToken[]): string | undefined {
  if (mecabTokens.length === 0) {
    return undefined;
  }

  const tokenStart = token.startPos ?? 0;
  const tokenEnd = token.endPos ?? tokenStart + token.surface.length;
  let bestSurfaceMatchPos1: string | undefined;
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
      bestSurfaceMatchPos1 = mecabToken.pos1;
    }
  }

  if (bestSurfaceMatchPos1) {
    return bestSurfaceMatchPos1;
  }

  let bestPos1: string | undefined;
  let bestOverlap = 0;
  let bestSpan = 0;
  let bestStartDistance = Number.MAX_SAFE_INTEGER;
  let bestStart = Number.MAX_SAFE_INTEGER;

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
      bestPos1 = mecabToken.pos1;
    }
  }

  return bestOverlap > 0 ? bestPos1 : undefined;
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

    let best: { pos1: string; index: number } | null = null;
    for (const candidate of indexedMecabTokens) {
      if (candidate.token.surface !== surface) {
        continue;
      }
      if (candidate.index < cursor) {
        continue;
      }
      best = { pos1: candidate.token.pos1 as string, index: candidate.index };
      break;
    }

    if (!best) {
      for (const candidate of indexedMecabTokens) {
        if (candidate.token.surface !== surface) {
          continue;
        }
        best = { pos1: candidate.token.pos1 as string, index: candidate.index };
        break;
      }
    }

    if (!best) {
      return token;
    }

    cursor = best.index + 1;
    return {
      ...token,
      pos1: best.pos1,
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

    const pos1 = pickClosestMecabPos1(token, mecabTokens);
    if (!pos1) {
      return token;
    }

    return {
      ...token,
      pos1,
    };
  });

  return fillMissingPos1BySurfaceSequence(overlapEnriched, mecabTokens);
}
