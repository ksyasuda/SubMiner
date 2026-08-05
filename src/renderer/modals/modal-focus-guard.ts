/**
 * Keeps focus inside an overlay-hosted modal.
 *
 * The overlay can lose focus to mpv or to the compositor while a modal is up,
 * which leaves the modal visible but inert. Recovery is debounced (and guarded
 * against re-entry) so a focus fight with the window manager cannot spin.
 */
export type ModalFocusGuardDeps = {
  isOpen: () => boolean;
  /** Modal root; focus inside it counts as "still in the modal". */
  getModalRoot: () => Element;
  /** Preferred focus targets in order; the first rendered one wins. */
  getPreferredFocusTargets: () => HTMLElement[];
  /** Used when no preferred target is rendered, e.g. the close button. */
  getFallbackFocusTarget: () => Element | null;
  /** Modal-layer windows own their focus; other layers ask the main window. */
  isModalLayer: boolean;
};

const FOCUS_RECOVERY_DEBOUNCE_MS = 120;

export function createModalFocusGuard(deps: ModalFocusGuardDeps) {
  let focusinGuard: ((event: FocusEvent) => void) | null = null;
  let windowFocusGuard: (() => void) | null = null;
  let pointerFocusGuard: ((event: Event) => void) | null = null;
  let isRecovering = false;
  let lastRecoveryAt = 0;

  function isModalFocusTarget(target: EventTarget | null): boolean {
    return target instanceof Element && deps.getModalRoot().contains(target);
  }

  function requestOverlayFocus(): void {
    if (!deps.isModalLayer) {
      // Best-effort: a rejected focus request must not surface as an unhandled
      // rejection, since this runs from blur/focus handlers.
      void Promise.resolve(window.electronAPI.focusMainWindow()).catch(() => {});
    }
  }

  function focusFallbackTarget(): boolean {
    requestOverlayFocus();

    // getClientRects() rather than offsetParent: the latter is null for
    // position:fixed elements, which would skip a perfectly visible target.
    const preferred = deps
      .getPreferredFocusTargets()
      .find((target) => target.getClientRects().length > 0);
    if (preferred) {
      preferred.focus({ preventScroll: true });
      return document.activeElement === preferred;
    }

    const fallback = deps.getFallbackFocusTarget();
    if (fallback instanceof HTMLElement) {
      fallback.focus({ preventScroll: true });
      return document.activeElement === fallback;
    }

    window.focus();
    return false;
  }

  function enforceModalFocus(): void {
    if (!deps.isOpen()) return;
    if (isModalFocusTarget(document.activeElement)) return;
    if (isRecovering) return;

    const now = Date.now();
    if (now - lastRecoveryAt < FOCUS_RECOVERY_DEBOUNCE_MS) return;

    isRecovering = true;
    lastRecoveryAt = now;
    focusFallbackTarget();
    window.setTimeout(() => {
      isRecovering = false;
    }, FOCUS_RECOVERY_DEBOUNCE_MS);
  }

  /** Idempotent; safe to call on every open. */
  function attach(): void {
    if (focusinGuard === null) {
      focusinGuard = (event: FocusEvent) => {
        if (!deps.isOpen()) return;
        if (!isModalFocusTarget(event.target)) {
          event.preventDefault();
          enforceModalFocus();
        }
      };
      document.addEventListener('focusin', focusinGuard);
    }

    if (pointerFocusGuard === null) {
      pointerFocusGuard = () => {
        requestOverlayFocus();
        enforceModalFocus();
      };
      const root = deps.getModalRoot();
      root.addEventListener('pointerdown', pointerFocusGuard);
      root.addEventListener('click', pointerFocusGuard);
    }

    if (windowFocusGuard === null) {
      windowFocusGuard = () => {
        requestOverlayFocus();
        enforceModalFocus();
      };
      window.addEventListener('blur', windowFocusGuard);
      window.addEventListener('focus', windowFocusGuard);
    }
  }

  function detach(): void {
    if (focusinGuard) {
      document.removeEventListener('focusin', focusinGuard);
      focusinGuard = null;
    }

    if (pointerFocusGuard) {
      const root = deps.getModalRoot();
      root.removeEventListener('pointerdown', pointerFocusGuard);
      root.removeEventListener('click', pointerFocusGuard);
      pointerFocusGuard = null;
    }

    if (windowFocusGuard) {
      window.removeEventListener('blur', windowFocusGuard);
      window.removeEventListener('focus', windowFocusGuard);
      windowFocusGuard = null;
    }
  }

  return {
    attach,
    detach,
    enforceModalFocus,
    focusFallbackTarget,
    isModalFocusTarget,
    requestOverlayFocus,
  };
}
