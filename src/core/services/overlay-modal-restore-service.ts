export function addOverlayModalRestoreFlagService<T extends string>(
  restoreSet: Set<T>,
  modal: T,
): void {
  restoreSet.add(modal);
}

export function handleOverlayModalClosedService<T extends string>(
  restoreSet: Set<T>,
  modal: T,
  setVisibleOverlayVisible: (visible: boolean) => void,
): void {
  if (!restoreSet.has(modal)) return;
  restoreSet.delete(modal);
  if (restoreSet.size === 0) {
    setVisibleOverlayVisible(false);
  }
}
