import type { ModalStateReader, RendererContext } from '../context';
import { syncOverlayMouseIgnoreState } from '../overlay-mouse-ignore.js';
import {
  YOMITAN_POPUP_HIDDEN_EVENT,
  YOMITAN_POPUP_SHOWN_EVENT,
  isYomitanPopupVisible,
  isYomitanPopupIframe,
} from '../yomitan-popup.js';

export function createMouseHandlers(
  ctx: RendererContext,
  options: {
    modalStateReader: ModalStateReader;
    applyYPercent: (yPercent: number) => void;
    getCurrentYPercent: () => number;
    persistSubtitlePositionPatch: (patch: { yPercent: number }) => void;
    getSubtitleHoverAutoPauseEnabled: () => boolean;
    getYomitanPopupAutoPauseEnabled: () => boolean;
    getPlaybackPaused: () => Promise<boolean | null>;
    sendMpvCommand: (command: (string | number)[]) => void;
  },
) {
  type HoverPointState = {
    overPrimarySubtitle: boolean;
    overSecondarySubtitle: boolean;
    isOverSubtitle: boolean;
  };

  let yomitanPopupVisible = false;
  let hoverPauseRequestId = 0;
  let popupPauseRequestId = 0;
  let pausedBySubtitleHover = false;
  let pausedByYomitanPopup = false;
  let lastPointerPosition: { clientX: number; clientY: number } | null = null;
  let pendingPointerResync = false;

  function reclaimOverlayWindowFocusForPopup(): void {
    if (!ctx.platform.shouldToggleMouseIgnore) {
      return;
    }
    if (ctx.platform.isMacOSPlatform || ctx.platform.isLinuxPlatform) {
      return;
    }

    if (typeof window.electronAPI.focusMainWindow === 'function') {
      void window.electronAPI.focusMainWindow();
    }
    window.focus();
    if (typeof ctx.dom.overlay.focus === 'function') {
      ctx.dom.overlay.focus({ preventScroll: true });
    }
  }

  function sustainPopupInteraction(): void {
    yomitanPopupVisible = true;
    ctx.state.yomitanPopupVisible = true;
    syncOverlayMouseIgnoreState(ctx);
  }

  function isElementWithinContainer(element: Element | null, container: HTMLElement): boolean {
    if (!element) {
      return false;
    }
    if (element === container) {
      return true;
    }
    return typeof container.contains === 'function' ? container.contains(element) : false;
  }

  function updatePointerPosition(event: MouseEvent | PointerEvent): void {
    lastPointerPosition = {
      clientX: event.clientX,
      clientY: event.clientY,
    };
  }

  function getHoverStateFromPoint(clientX: number, clientY: number): HoverPointState {
    const hoveredElement =
      typeof document.elementFromPoint === 'function'
        ? document.elementFromPoint(clientX, clientY)
        : null;
    const overPrimarySubtitle = isElementWithinContainer(hoveredElement, ctx.dom.subtitleContainer);
    const overSecondarySubtitle = isElementWithinContainer(
      hoveredElement,
      ctx.dom.secondarySubContainer,
    );

    return {
      overPrimarySubtitle,
      overSecondarySubtitle,
      isOverSubtitle: overPrimarySubtitle || overSecondarySubtitle,
    };
  }

  function syncHoverStateFromPoint(clientX: number, clientY: number): HoverPointState {
    const hoverState = getHoverStateFromPoint(clientX, clientY);

    ctx.state.isOverSubtitle = hoverState.isOverSubtitle;
    ctx.dom.secondarySubContainer.classList.toggle(
      'secondary-sub-hover-active',
      hoverState.overSecondarySubtitle,
    );

    return hoverState;
  }

  function syncHoverStateFromTrackedPointer(event: MouseEvent | PointerEvent): void {
    if (!ctx.platform.shouldToggleMouseIgnore || ctx.state.isDragging) {
      return;
    }

    const wasOverSubtitle = ctx.state.isOverSubtitle;
    const wasOverSecondarySubtitle = ctx.dom.secondarySubContainer.classList.contains(
      'secondary-sub-hover-active',
    );
    const hoverState = syncHoverStateFromPoint(event.clientX, event.clientY);

    if (!wasOverSubtitle && hoverState.isOverSubtitle) {
      void handleMouseEnter(undefined, hoverState.overSecondarySubtitle);
      return;
    }

    if (wasOverSubtitle && !hoverState.isOverSubtitle) {
      void handleMouseLeave(undefined, wasOverSecondarySubtitle);
      return;
    }

    if (
      hoverState.isOverSubtitle &&
      hoverState.overSecondarySubtitle !== wasOverSecondarySubtitle
    ) {
      syncOverlayMouseIgnoreState(ctx);
    }
  }

  function restorePointerInteractionState(): void {
    const pointerPosition = lastPointerPosition;
    pendingPointerResync = false;
    if (pointerPosition) {
      syncHoverStateFromPoint(pointerPosition.clientX, pointerPosition.clientY);
    } else {
      ctx.state.isOverSubtitle = false;
      ctx.dom.secondarySubContainer.classList.remove('secondary-sub-hover-active');
    }
    syncOverlayMouseIgnoreState(ctx);

    if (!ctx.platform.shouldToggleMouseIgnore || ctx.state.isOverSubtitle) {
      return;
    }

    pendingPointerResync = true;
    ctx.dom.overlay.classList.add('interactive');
    window.electronAPI.setIgnoreMouseEvents(false);
  }

  function maybeResyncPointerHoverState(event: MouseEvent | PointerEvent): void {
    if (!pendingPointerResync) {
      return;
    }
    pendingPointerResync = false;
    syncHoverStateFromPoint(event.clientX, event.clientY);
    syncOverlayMouseIgnoreState(ctx);
  }

  function isWithinOtherSubtitleContainer(
    relatedTarget: EventTarget | null,
    otherContainer: HTMLElement,
  ): boolean {
    if (relatedTarget === otherContainer) {
      return true;
    }
    if (typeof Node !== 'undefined' && relatedTarget instanceof Node) {
      return otherContainer.contains(relatedTarget);
    }
    return false;
  }

  function maybeResumeHoverPause(): void {
    if (!pausedBySubtitleHover) return;
    if (pausedByYomitanPopup) return;
    if (ctx.state.isOverSubtitle) return;
    pausedBySubtitleHover = false;
    options.sendMpvCommand(['set_property', 'pause', 'no']);
  }

  function maybeResumeYomitanPopupPause(): void {
    if (!pausedByYomitanPopup) return;
    pausedByYomitanPopup = false;
    if (ctx.state.isOverSubtitle && options.getSubtitleHoverAutoPauseEnabled()) {
      pausedBySubtitleHover = true;
      return;
    }
    options.sendMpvCommand(['set_property', 'pause', 'no']);
  }

  async function maybePauseForYomitanPopup(): Promise<void> {
    if (!yomitanPopupVisible || !options.getYomitanPopupAutoPauseEnabled()) {
      return;
    }

    const requestId = ++popupPauseRequestId;
    if (pausedByYomitanPopup) return;

    if (pausedBySubtitleHover) {
      pausedBySubtitleHover = false;
      pausedByYomitanPopup = true;
      return;
    }

    let paused: boolean | null = null;
    try {
      paused = await options.getPlaybackPaused();
    } catch {
      return;
    }

    if (
      requestId !== popupPauseRequestId ||
      !yomitanPopupVisible ||
      !options.getYomitanPopupAutoPauseEnabled()
    ) {
      return;
    }
    if (paused !== false) return;

    options.sendMpvCommand(['set_property', 'pause', 'yes']);
    pausedByYomitanPopup = true;
  }

  function enablePopupInteraction(): void {
    sustainPopupInteraction();
    if (ctx.platform.isMacOSPlatform) {
      window.focus();
    }
  }

  function disablePopupInteractionIfIdle(): void {
    if (typeof document !== 'undefined' && isYomitanPopupVisible(document)) {
      sustainPopupInteraction();
      reclaimOverlayWindowFocusForPopup();
      return;
    }

    yomitanPopupVisible = false;
    ctx.state.yomitanPopupVisible = false;
    popupPauseRequestId += 1;
    maybeResumeYomitanPopupPause();
    maybeResumeHoverPause();
    syncOverlayMouseIgnoreState(ctx);
  }

  async function handleMouseEnter(_event?: MouseEvent, showSecondaryHover = false): Promise<void> {
    ctx.state.isOverSubtitle = true;
    if (showSecondaryHover) {
      ctx.dom.secondarySubContainer.classList.add('secondary-sub-hover-active');
    }
    syncOverlayMouseIgnoreState(ctx);

    if (yomitanPopupVisible && options.getYomitanPopupAutoPauseEnabled()) {
      return;
    }

    if (!options.getSubtitleHoverAutoPauseEnabled()) {
      return;
    }

    if (pausedBySubtitleHover) {
      return;
    }

    const requestId = ++hoverPauseRequestId;
    let paused: boolean | null = null;
    try {
      paused = await options.getPlaybackPaused();
    } catch {
      return;
    }
    if (requestId !== hoverPauseRequestId || !ctx.state.isOverSubtitle) {
      return;
    }
    if (paused !== false) {
      return;
    }
    options.sendMpvCommand(['set_property', 'pause', 'yes']);
    pausedBySubtitleHover = true;
  }

  async function handleMouseLeave(_event?: MouseEvent, hideSecondaryHover = false): Promise<void> {
    const relatedTarget = _event?.relatedTarget ?? null;
    const otherContainer = hideSecondaryHover
      ? ctx.dom.subtitleContainer
      : ctx.dom.secondarySubContainer;
    if (relatedTarget && isWithinOtherSubtitleContainer(relatedTarget, otherContainer)) {
      ctx.state.isOverSubtitle = false;
      if (hideSecondaryHover) {
        ctx.dom.secondarySubContainer.classList.remove('secondary-sub-hover-active');
      }
      return;
    }

    ctx.state.isOverSubtitle = false;
    if (hideSecondaryHover) {
      ctx.dom.secondarySubContainer.classList.remove('secondary-sub-hover-active');
    }
    hoverPauseRequestId += 1;
    maybeResumeHoverPause();
    if (yomitanPopupVisible) return;
    disablePopupInteractionIfIdle();
  }

  function setupDragging(): void {
    ctx.dom.subtitleContainer.addEventListener('mousedown', (e: MouseEvent) => {
      if (e.button === 2) {
        e.preventDefault();
        ctx.state.isDragging = true;
        ctx.state.dragStartY = e.clientY;
        ctx.state.startYPercent = options.getCurrentYPercent();
        ctx.dom.subtitleContainer.style.cursor = 'grabbing';
      }
    });

    document.addEventListener('mousemove', (e: MouseEvent) => {
      if (!ctx.state.isDragging) return;

      const deltaY = ctx.state.dragStartY - e.clientY;
      const deltaPercent = (deltaY / window.innerHeight) * 100;
      const newYPercent = ctx.state.startYPercent + deltaPercent;

      options.applyYPercent(newYPercent);
    });

    document.addEventListener('mouseup', (e: MouseEvent) => {
      if (ctx.state.isDragging && e.button === 2) {
        ctx.state.isDragging = false;
        ctx.dom.subtitleContainer.style.cursor = '';

        const yPercent = options.getCurrentYPercent();
        options.persistSubtitlePositionPatch({ yPercent });
      }
    });

    ctx.dom.subtitleContainer.addEventListener('contextmenu', (e: Event) => {
      e.preventDefault();
    });
  }

  function setupResizeHandler(): void {
    window.addEventListener('resize', () => {
      options.applyYPercent(options.getCurrentYPercent());
    });
  }

  function setupPointerTracking(): void {
    document.addEventListener('mousemove', (event: MouseEvent) => {
      updatePointerPosition(event);
      syncHoverStateFromTrackedPointer(event);
      maybeResyncPointerHoverState(event);
    });
    document.addEventListener('pointermove', (event: PointerEvent) => {
      updatePointerPosition(event);
      syncHoverStateFromTrackedPointer(event);
      maybeResyncPointerHoverState(event);
    });
  }

  function setupSelectionObserver(): void {
    document.addEventListener('selectionchange', () => {
      const selection = window.getSelection();
      const hasSelection = selection && selection.rangeCount > 0 && !selection.isCollapsed;

      if (hasSelection) {
        ctx.dom.subtitleRoot.classList.add('has-selection');
      } else {
        ctx.dom.subtitleRoot.classList.remove('has-selection');
      }
    });
  }

  function setupYomitanObserver(): void {
    yomitanPopupVisible = isYomitanPopupVisible(document);
    ctx.state.yomitanPopupVisible = yomitanPopupVisible;
    void maybePauseForYomitanPopup();

    window.addEventListener(YOMITAN_POPUP_SHOWN_EVENT, () => {
      enablePopupInteraction();
      void maybePauseForYomitanPopup();
    });

    window.addEventListener(YOMITAN_POPUP_HIDDEN_EVENT, () => {
      disablePopupInteractionIfIdle();
    });

    const observer = new MutationObserver((mutations: MutationRecord[]) => {
      for (const mutation of mutations) {
        mutation.addedNodes.forEach((node) => {
          if (node.nodeType !== Node.ELEMENT_NODE) return;
          const element = node as Element;
          if (isYomitanPopupIframe(element)) {
            enablePopupInteraction();
            void maybePauseForYomitanPopup();
          }
        });

        mutation.removedNodes.forEach((node) => {
          if (node.nodeType !== Node.ELEMENT_NODE) return;
          const element = node as Element;
          if (isYomitanPopupIframe(element)) {
            disablePopupInteractionIfIdle();
          }
        });
      }
    });

    observer.observe(document.body, {
      childList: true,
      subtree: true,
    });
  }

  return {
    handlePrimaryMouseEnter: (event?: MouseEvent) => handleMouseEnter(event, false),
    handlePrimaryMouseLeave: (event?: MouseEvent) => handleMouseLeave(event, false),
    handleSecondaryMouseEnter: (event?: MouseEvent) => handleMouseEnter(event, true),
    handleSecondaryMouseLeave: (event?: MouseEvent) => handleMouseLeave(event, true),
    handleMouseEnter,
    handleMouseLeave,
    restorePointerInteractionState,
    setupDragging,
    setupPointerTracking,
    setupResizeHandler,
    setupSelectionObserver,
    setupYomitanObserver,
  };
}
