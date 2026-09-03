/** Repository version codes are monotonic, so only a strictly newer code is an update. */
export function hasExtensionUpdate(
  installedVersionCode: number | null,
  availableVersionCode: number,
): boolean {
  return installedVersionCode !== null && availableVersionCode > installedVersionCode;
}

/** Keep only installed packages whose repository build is strictly newer. */
export function findExtensionUpdates<T extends { pkg: string; versionCode: number }>(
  installed: ReadonlyArray<{ pkg: string; versionCode: number | null }>,
  offered: readonly T[],
): T[] {
  const installedVersions = new Map(
    installed.map((extension) => [extension.pkg, extension.versionCode]),
  );
  return offered.filter((extension) => {
    const installedVersion = installedVersions.get(extension.pkg);
    return installedVersion === undefined
      ? false
      : hasExtensionUpdate(installedVersion, extension.versionCode);
  });
}
