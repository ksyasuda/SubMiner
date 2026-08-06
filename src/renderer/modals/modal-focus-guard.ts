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
  let pointerFocusRoot: Element | null = null;
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
    // Rendered is not the same as focusable, so keep trying until one sticks
    // instead of giving up on the first candidate that refuses focus.
    for (const target of deps.getPreferredFocusTargets()) {
      if (target.getClientRects().length === 0) continue;
      target.focus({ preventScroll: true });
      if (document.activeElement === target) return true;
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
      // focusin is not cancelable, so there is nothing to preventDefault here;
      // focus is taken back afterwards instead.
      focusinGuard = (event: FocusEvent) => {
        if (!deps.isOpen()) return;
        if (!isModalFocusTarget(event.target)) {
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
      // Remember the root we bound to: resolving it again on detach could
      // return a different element and leak the listeners on the old one.
      pointerFocusRoot = deps.getModalRoot();
      pointerFocusRoot.addEventListener('pointerdown', pointerFocusGuard);
      pointerFocusRoot.addEventListener('click', pointerFocusGuard);
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
      pointerFocusRoot?.removeEventListener('pointerdown', pointerFocusGuard);
      pointerFocusRoot?.removeEventListener('click', pointerFocusGuard);
      pointerFocusGuard = null;
      pointerFocusRoot = null;
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
