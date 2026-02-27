import type { ModalStateReader, RendererContext } from '../context';
import {
  YOMITAN_POPUP_HIDDEN_EVENT,
  YOMITAN_POPUP_SHOWN_EVENT,
  hasYomitanPopupIframe,
  isYomitanPopupIframe,
} from '../yomitan-popup.js';

export function createMouseHandlers(
  ctx: RendererContext,
  options: {
    modalStateReader: ModalStateReader;
    applyYPercent: (yPercent: number) => void;
    getCurrentYPercent: () => number;
    persistSubtitlePositionPatch: (patch: { yPercent: number }) => void;
  },
) {
  let yomitanPopupVisible = false;

  function enablePopupInteraction(): void {
    yomitanPopupVisible = true;
    ctx.dom.overlay.classList.add('interactive');
    if (ctx.platform.shouldToggleMouseIgnore) {
      window.electronAPI.setIgnoreMouseEvents(false);
    }
    if (ctx.platform.isMacOSPlatform) {
      window.focus();
    }
  }

  function disablePopupInteractionIfIdle(): void {
    if (hasYomitanPopupIframe(document)) {
      yomitanPopupVisible = true;
      return;
    }

    yomitanPopupVisible = false;
    if (
      !ctx.state.isOverSubtitle &&
      !options.modalStateReader.isAnyModalOpen()
    ) {
      ctx.dom.overlay.classList.remove('interactive');
      if (ctx.platform.shouldToggleMouseIgnore) {
        window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
      }
    }
  }

  function handleMouseEnter(): void {
    ctx.state.isOverSubtitle = true;
    ctx.dom.overlay.classList.add('interactive');
    if (ctx.platform.shouldToggleMouseIgnore) {
      window.electronAPI.setIgnoreMouseEvents(false);
    }
  }

  function handleMouseLeave(): void {
    ctx.state.isOverSubtitle = false;
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
    yomitanPopupVisible = hasYomitanPopupIframe(document);

    window.addEventListener(YOMITAN_POPUP_SHOWN_EVENT, () => {
      enablePopupInteraction();
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
    handleMouseEnter,
    handleMouseLeave,
    setupDragging,
    setupResizeHandler,
    setupSelectionObserver,
    setupYomitanObserver,
  };
}
