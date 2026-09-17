// macOS-only: Electron's setVisibleOnAllWorkspaces(..., { visibleOnFullScreen: true })
// transforms the whole app into a UIElement (accessory) process so overlay windows can
// float above fullscreen Spaces. That side effect removes the Dock icon and the app's
// Cmd+Tab entry, making regular windows (anime browser, settings) unreachable.
//
// While at least one regular window "retains" the Dock icon, overlay/stats level
// assertions must pass skipTransformProcessType so they don't yank the app back to
// accessory mode behind the user's back.

type DockLike = {
  show: () => Promise<void> | void;
  hide: () => void;
};

let dockIconRetainCount = 0;

export function isDockIconRetained(): boolean {
  return dockIconRetainCount > 0;
}

export function retainDockIcon(options: {
  dock: DockLike | null | undefined;
  platform?: NodeJS.Platform;
}): void {
  if ((options.platform ?? process.platform) !== 'darwin') return;
  dockIconRetainCount += 1;
  if (dockIconRetainCount === 1) {
    void options.dock?.show();
  }
}

export function releaseDockIcon(options: {
  dock: DockLike | null | undefined;
  platform?: NodeJS.Platform;
  // True when an overlay window still needs the accessory transform to float
  // above fullscreen Spaces; false leaves the Dock icon visible.
  shouldRehide: () => boolean;
}): void {
  if ((options.platform ?? process.platform) !== 'darwin') return;
  if (dockIconRetainCount === 0) return;
  dockIconRetainCount -= 1;
  if (dockIconRetainCount === 0 && options.shouldRehide()) {
    options.dock?.hide();
  }
}

export function resetDockIconRetentionForTests(): void {
  dockIconRetainCount = 0;
}
