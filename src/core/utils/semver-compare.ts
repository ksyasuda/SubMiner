/**
 * Loose semver ordering shared by the updater and the changelog UI.
 * Returns >0 when `a` is newer, <0 when older, 0 when equal.
 */
export function compareSemverLike(a: string, b: string): number {
  const parse = (
    value: string,
  ): {
    core: number[];
    prerelease: Array<number | string>;
  } => {
    // Build metadata ("+build.2") is not part of precedence per semver, and
    // leaving it attached makes it leak into the prerelease comparison.
    const normalized = value.replace(/^v/i, '').split('+', 1)[0] ?? '';
    const [coreText = '', ...prereleaseParts] = normalized.split('-');
    const core = coreText
      .split('.')
      .slice(0, 3)
      .map((part) => Number.parseInt(part, 10) || 0);
    while (core.length < 3) core.push(0);
    const prereleaseText = prereleaseParts.join('-');
    return {
      core,
      prerelease: prereleaseText
        ? prereleaseText.split('.').map((part) => {
            const numeric = Number.parseInt(part, 10);
            return /^\d+$/.test(part) ? numeric : part;
          })
        : [],
    };
  };
  const left = parse(a);
  const right = parse(b);
  for (let i = 0; i < 3; i += 1) {
    const diff = (left.core[i] ?? 0) - (right.core[i] ?? 0);
    if (diff !== 0) return diff;
  }

  if (left.prerelease.length === 0 && right.prerelease.length === 0) return 0;
  if (left.prerelease.length === 0) return 1;
  if (right.prerelease.length === 0) return -1;

  const length = Math.max(left.prerelease.length, right.prerelease.length);
  for (let i = 0; i < length; i += 1) {
    const leftPart = left.prerelease[i];
    const rightPart = right.prerelease[i];
    if (leftPart === undefined && rightPart === undefined) return 0;
    if (leftPart === undefined) return -1;
    if (rightPart === undefined) return 1;
    if (leftPart === rightPart) continue;
    if (typeof leftPart === 'number' && typeof rightPart === 'number') {
      return leftPart - rightPart;
    }
    if (typeof leftPart === 'number') return -1;
    if (typeof rightPart === 'number') return 1;
    return leftPart > rightPart ? 1 : -1;
  }
  return 0;
}
