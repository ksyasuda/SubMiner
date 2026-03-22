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
  let yomitanPopupVisible = false;
  let hoverPauseRequestId = 0;
  let popupPauseRequestId = 0;
  let pausedBySubtitleHover = false;
  let pausedByYomitanPopup = false;

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
    yomitanPopupVisible = true;
    ctx.state.yomitanPopupVisible = true;
    syncOverlayMouseIgnoreState(ctx);
    if (ctx.platform.isMacOSPlatform) {
      window.focus();
    }
  }

  function disablePopupInteractionIfIdle(): void {
    if (typeof document !== 'undefined' && isYomitanPopupVisible(document)) {
      yomitanPopupVisible = true;
      ctx.state.yomitanPopupVisible = true;
      return;
    }

    yomitanPopupVisible = false;
    ctx.state.yomitanPopupVisible = false;
    popupPauseRequestId += 1;
    maybeResumeYomitanPopupPause();
    maybeResumeHoverPause();
    syncOverlayMouseIgnoreState(ctx);
  }

  async function handleMouseEnter(
    _event?: MouseEvent,
    showSecondaryHover = false,
  ): Promise<void> {
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

  async function handleMouseLeave(
    _event?: MouseEvent,
    hideSecondaryHover = false,
  ): Promise<void> {
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
    setupDragging,
    setupResizeHandler,
    setupSelectionObserver,
    setupYomitanObserver,
  };
}
