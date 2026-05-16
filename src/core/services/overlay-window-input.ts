export type OverlayWindowKind = 'visible' | 'modal';

export function isTabInputForMpvForwarding(input: Electron.Input): boolean {
  if (input.type !== 'keyDown' || input.isAutoRepeat) return false;
  if (input.alt || input.control || input.meta || input.shift) return false;
  return input.code === 'Tab' || input.key === 'Tab';
}

function isLookupWindowToggleInput(input: Electron.Input): boolean {
  if (input.type !== 'keyDown') return false;
  if (input.alt) return false;
  if (!input.control && !input.meta) return false;
  if (input.shift) return false;
  const normalizedKey = typeof input.key === 'string' ? input.key.toLowerCase() : '';
  return input.code === 'KeyY' || normalizedKey === 'y';
}

function isKeyboardModeToggleInput(input: Electron.Input): boolean {
  if (input.type !== 'keyDown') return false;
  if (input.alt) return false;
  if (!input.control && !input.meta) return false;
  if (!input.shift) return false;
  const normalizedKey = typeof input.key === 'string' ? input.key.toLowerCase() : '';
  return input.code === 'KeyY' || normalizedKey === 'y';
}

export function handleOverlayWindowBeforeInputEvent(options: {
  kind: OverlayWindowKind;
  windowVisible: boolean;
  input: Electron.Input;
  preventDefault: () => void;
  sendKeyboardModeToggleRequested: () => void;
  sendLookupWindowToggleRequested: () => void;
  tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
  forwardTabToMpv: () => void;
}): boolean {
  if (options.kind === 'modal') return false;
  if (!options.windowVisible) return false;

  if (isKeyboardModeToggleInput(options.input)) {
    options.preventDefault();
    options.sendKeyboardModeToggleRequested();
    return true;
  }

  if (isLookupWindowToggleInput(options.input)) {
    options.preventDefault();
    options.sendLookupWindowToggleRequested();
    return true;
  }

  if (isTabInputForMpvForwarding(options.input)) {
    options.preventDefault();
    options.forwardTabToMpv();
    return true;
  }

  if (!options.tryHandleOverlayShortcutLocalFallback(options.input)) return false;
  options.preventDefault();
  return true;
}

export function handleOverlayWindowBlurred(options: {
  kind: OverlayWindowKind;
  windowVisible: boolean;
  isOverlayVisible: (kind: OverlayWindowKind) => boolean;
  ensureOverlayWindowLevel: () => void;
  moveWindowTop: () => void;
  onVisibleOverlayBlur?: () => void;
  platform?: NodeJS.Platform;
}): boolean {
  const platform = options.platform ?? process.platform;
  if (platform === 'win32' && options.kind === 'visible') {
    options.onVisibleOverlayBlur?.();
    return false;
  }
  if (platform === 'darwin' && options.kind === 'visible') {
    options.onVisibleOverlayBlur?.();
    return false;
  }

  if (options.kind === 'visible' && !options.isOverlayVisible(options.kind)) {
    return false;
  }

  options.ensureOverlayWindowLevel();
  if (options.kind === 'visible' && options.windowVisible) {
    options.moveWindowTop();
  }
  return true;
}
