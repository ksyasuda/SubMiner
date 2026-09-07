import type { RendererContext } from '../context';
import { syncOverlayMouseIgnoreState } from '../overlay-mouse-ignore';

function isEditingText(target: EventTarget | null): boolean {
  return (
    target instanceof HTMLElement &&
    (target.isContentEditable ||
      target instanceof HTMLInputElement ||
      target instanceof HTMLTextAreaElement)
  );
}

function getSelectionRange(list: HTMLElement): Range | null {
  const selection = list.ownerDocument?.defaultView?.getSelection();
  if (!selection || selection.isCollapsed || selection.rangeCount === 0) return null;
  const range = selection.getRangeAt(0);
  return list.contains(range.startContainer) && list.contains(range.endContainer) ? range : null;
}

export function hasSubtitleSidebarSelection(list: HTMLElement): boolean {
  return getSelectionRange(list) !== null;
}

// Read only dialogue nodes, preserving partial first/last lines and DOM cue order.
export function getSubtitleSidebarSelection(list: HTMLElement): {
  text: string;
  cueCount: number;
} | null {
  const range = getSelectionRange(list);
  if (!range) return null;

  const parts: string[] = [];
  for (const text of list.querySelectorAll<HTMLElement>('.subtitle-sidebar-text')) {
    if (!range.intersectsNode(text)) continue;
    const part = list.ownerDocument.createRange();
    part.selectNodeContents(text);
    if (range.compareBoundaryPoints(Range.START_TO_START, part) > 0) {
      part.setStart(range.startContainer, range.startOffset);
    }
    if (range.compareBoundaryPoints(Range.END_TO_END, part) < 0) {
      part.setEnd(range.endContainer, range.endOffset);
    }
    const selectedText = part.toString();
    if (selectedText.trim()) parts.push(selectedText);
  }
  return parts.length > 0 ? { text: parts.join('\n\n'), cueCount: parts.length } : null;
}

export function clearSubtitleSidebarSelection(list: HTMLElement): void {
  const selection = list.ownerDocument?.defaultView?.getSelection();
  if (selection?.anchorNode && list.contains(selection.anchorNode)) {
    selection.removeAllRanges();
  }
}

export function wireSubtitleSidebarSelection(ctx: RendererContext): () => void {
  const list = ctx.dom.subtitleSidebarList;
  const button = ctx.dom.subtitleSidebarCopy;
  const doc = list.ownerDocument;
  const abort = new AbortController();
  const { signal } = abort;

  const updateButton = () => {
    const selected = getSubtitleSidebarSelection(list);
    button.hidden = !selected;
    button.textContent = selected
      ? `Copy ${selected.cueCount} ${selected.cueCount === 1 ? 'line' : 'lines'}`
      : 'Copy';
  };
  const copied = () => {
    ctx.dom.subtitleSidebarStatus.textContent = 'Selection copied.';
  };
  const copySelection = async () => {
    const selected = getSubtitleSidebarSelection(list);
    if (!selected) return;
    try {
      await window.electronAPI.copySubtitleSidebarSelection(selected.text);
      copied();
    } catch {
      ctx.dom.subtitleSidebarStatus.textContent = 'Could not copy selection. Try again.';
    }
  };

  doc.addEventListener('selectionchange', updateButton, { signal });
  doc.addEventListener(
    'copy',
    (event) => {
      if (!ctx.state.subtitleSidebarModalOpen || isEditingText(event.target)) return;
      const selected = getSubtitleSidebarSelection(list);
      if (!selected || !event.clipboardData) return;
      event.preventDefault();
      event.clipboardData.setData('text/plain', selected.text);
      copied();
    },
    { signal },
  );
  // Capture before modal Escape handling and the current-subtitle copy shortcut.
  doc.addEventListener(
    'keydown',
    (event) => {
      if (
        !ctx.state.subtitleSidebarModalOpen ||
        isEditingText(event.target) ||
        !getSubtitleSidebarSelection(list)
      )
        return;
      if (event.key === 'Escape') {
        event.preventDefault();
        event.stopImmediatePropagation();
        clearSubtitleSidebarSelection(list);
        updateButton();
      } else if (
        (event.ctrlKey || event.metaKey) &&
        !event.altKey &&
        !event.shiftKey &&
        event.key.toLowerCase() === 'c'
      ) {
        event.preventDefault();
        event.stopImmediatePropagation();
        void copySelection();
      }
    },
    { capture: true, signal },
  );
  button.addEventListener('mousedown', (event) => event.preventDefault(), { signal });
  button.addEventListener('click', copySelection, { signal });
  list.addEventListener(
    'pointerdown',
    (event) => {
      if (event.button === 0) {
        list.dataset.selecting = 'true';
        list.scrollTo({ top: list.scrollTop, behavior: 'instant' });
        syncOverlayMouseIgnoreState(ctx);
      }
    },
    { signal },
  );
  const stopDragging = () => {
    delete list.dataset.selecting;
    syncOverlayMouseIgnoreState(ctx);
  };
  doc.addEventListener('pointerup', stopDragging, { signal });
  doc.addEventListener('pointercancel', stopDragging, { signal });
  doc.defaultView?.addEventListener('blur', stopDragging, { signal });
  updateButton();
  return () => {
    abort.abort();
    stopDragging();
  };
}
