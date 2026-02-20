import type { ModalStateReader, RendererContext } from '../context';

export function createMouseHandlers(
  ctx: RendererContext,
  options: {
    modalStateReader: ModalStateReader;
    applyInvisibleSubtitleLayoutFromMpvMetrics: (metrics: any, source: string) => void;
    applyYPercent: (yPercent: number) => void;
    getCurrentYPercent: () => number;
    persistSubtitlePositionPatch: (patch: { yPercent: number }) => void;
  },
) {
  const wordSegmenter =
    typeof Intl !== 'undefined' && 'Segmenter' in Intl
      ? new Intl.Segmenter(undefined, { granularity: 'word' })
      : null;

  function handleMouseEnter(): void {
    ctx.state.isOverSubtitle = true;
    ctx.dom.overlay.classList.add('interactive');
    if (ctx.platform.shouldToggleMouseIgnore) {
      window.electronAPI.setIgnoreMouseEvents(false);
    }
  }

  function handleMouseLeave(): void {
    ctx.state.isOverSubtitle = false;
    const yomitanPopup = document.querySelector('iframe[id^="yomitan-popup"]');
    if (
      !yomitanPopup &&
      !options.modalStateReader.isAnyModalOpen() &&
      !ctx.state.invisiblePositionEditMode
    ) {
      ctx.dom.overlay.classList.remove('interactive');
      if (ctx.platform.shouldToggleMouseIgnore) {
        window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
      }
    }
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

  function getCaretTextPointRange(clientX: number, clientY: number): Range | null {
    const documentWithCaretApi = document as Document & {
      caretRangeFromPoint?: (x: number, y: number) => Range | null;
      caretPositionFromPoint?: (
        x: number,
        y: number,
      ) => { offsetNode: Node; offset: number } | null;
    };

    if (typeof documentWithCaretApi.caretRangeFromPoint === 'function') {
      return documentWithCaretApi.caretRangeFromPoint(clientX, clientY);
    }

    if (typeof documentWithCaretApi.caretPositionFromPoint === 'function') {
      const caretPosition = documentWithCaretApi.caretPositionFromPoint(clientX, clientY);
      if (!caretPosition) return null;
      const range = document.createRange();
      range.setStart(caretPosition.offsetNode, caretPosition.offset);
      range.collapse(true);
      return range;
    }

    return null;
  }

  function getWordBoundsAtOffset(
    text: string,
    offset: number,
  ): { start: number; end: number } | null {
    if (!text || text.length === 0) return null;

    const clampedOffset = Math.max(0, Math.min(offset, text.length));
    const probeIndex = clampedOffset >= text.length ? Math.max(0, text.length - 1) : clampedOffset;

    if (wordSegmenter) {
      for (const part of wordSegmenter.segment(text)) {
        const start = part.index;
        const end = start + part.segment.length;
        if (probeIndex >= start && probeIndex < end) {
          if (part.isWordLike === false) return null;
          return { start, end };
        }
      }
    }

    const isBoundary = (char: string): boolean =>
      /[\s\u3000.,!?;:()[\]{}"'`~<>/\\|@#$%^&*+=\-、。・「」『』【】〈〉《》]/.test(char);

    const probeChar = text[probeIndex];
    if (!probeChar || isBoundary(probeChar)) return null;

    let start = probeIndex;
    while (start > 0 && !isBoundary(text[start - 1] ?? '')) {
      start -= 1;
    }

    let end = probeIndex + 1;
    while (end < text.length && !isBoundary(text[end] ?? '')) {
      end += 1;
    }

    if (end <= start) return null;
    return { start, end };
  }

  function updateHoverWordSelection(event: MouseEvent): void {
    if (!ctx.platform.isInvisibleLayer || !ctx.platform.isMacOSPlatform) return;
    if (event.buttons !== 0) return;
    if (!(event.target instanceof Node)) return;
    if (!ctx.dom.subtitleRoot.contains(event.target)) return;

    const caretRange = getCaretTextPointRange(event.clientX, event.clientY);
    if (!caretRange) return;
    if (caretRange.startContainer.nodeType !== Node.TEXT_NODE) return;
    if (!ctx.dom.subtitleRoot.contains(caretRange.startContainer)) return;

    const textNode = caretRange.startContainer as Text;
    const wordBounds = getWordBoundsAtOffset(textNode.data, caretRange.startOffset);
    if (!wordBounds) return;

    const selectionKey = `${wordBounds.start}:${wordBounds.end}:${textNode.data.slice(
      wordBounds.start,
      wordBounds.end,
    )}`;
    if (
      selectionKey === ctx.state.lastHoverSelectionKey &&
      textNode === ctx.state.lastHoverSelectionNode
    ) {
      return;
    }

    const selection = window.getSelection();
    if (!selection) return;

    const range = document.createRange();
    range.setStart(textNode, wordBounds.start);
    range.setEnd(textNode, wordBounds.end);

    selection.removeAllRanges();
    selection.addRange(range);
    ctx.state.lastHoverSelectionKey = selectionKey;
    ctx.state.lastHoverSelectionNode = textNode;
  }

  function setupInvisibleHoverSelection(): void {
    if (!ctx.platform.isInvisibleLayer || !ctx.platform.isMacOSPlatform) return;

    ctx.dom.subtitleRoot.addEventListener('mousemove', (event: MouseEvent) => {
      updateHoverWordSelection(event);
    });

    ctx.dom.subtitleRoot.addEventListener('mouseleave', () => {
      ctx.state.lastHoverSelectionKey = '';
      ctx.state.lastHoverSelectionNode = null;
    });
  }

  function setupResizeHandler(): void {
    window.addEventListener('resize', () => {
      if (ctx.platform.isInvisibleLayer) {
        if (!ctx.state.mpvSubtitleRenderMetrics) return;
        options.applyInvisibleSubtitleLayoutFromMpvMetrics(
          ctx.state.mpvSubtitleRenderMetrics,
          'resize',
        );
        return;
      }
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
    const observer = new MutationObserver((mutations: MutationRecord[]) => {
      for (const mutation of mutations) {
        mutation.addedNodes.forEach((node) => {
          if (node.nodeType !== Node.ELEMENT_NODE) return;
          const element = node as Element;
          if (
            element.tagName === 'IFRAME' &&
            element.id &&
            element.id.startsWith('yomitan-popup')
          ) {
            ctx.dom.overlay.classList.add('interactive');
            if (ctx.platform.shouldToggleMouseIgnore) {
              window.electronAPI.setIgnoreMouseEvents(false);
            }
          }
        });

        mutation.removedNodes.forEach((node) => {
          if (node.nodeType !== Node.ELEMENT_NODE) return;
          const element = node as Element;
          if (
            element.tagName === 'IFRAME' &&
            element.id &&
            element.id.startsWith('yomitan-popup')
          ) {
            if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
              ctx.dom.overlay.classList.remove('interactive');
              if (ctx.platform.shouldToggleMouseIgnore) {
                window.electronAPI.setIgnoreMouseEvents(true, {
                  forward: true,
                });
              }
            }
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
    setupInvisibleHoverSelection,
    setupResizeHandler,
    setupSelectionObserver,
    setupYomitanObserver,
  };
}
