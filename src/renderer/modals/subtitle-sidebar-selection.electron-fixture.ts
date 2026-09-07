import type { ElectronAPI, SubtitleSidebarSnapshot } from '../../types';
import { SUBTITLE_DEFAULT_CONFIG } from '../../config/definitions/defaults-subtitle';
import { createRendererState } from '../state';
import { resolveRendererDom } from '../utils/dom';
import { resolvePlatformInfo } from '../utils/platform';
import { createSubtitleSidebarModal } from './subtitle-sidebar';
import {
  getSubtitleSidebarSelection,
  wireSubtitleSidebarSelection,
} from './subtitle-sidebar-selection';

export async function setup() {
  const commands: unknown[] = [];
  const snapshot: SubtitleSidebarSnapshot = {
    sourceKey: 'episode-1:track-1',
    cues: [
      { text: '最初の台詞', startTime: 0, endTime: 1 },
      { text: '同じ台詞\n二行目', startTime: 1, endTime: 2 },
      { text: '同じ台詞', startTime: 2, endTime: 3 },
      ...Array.from({ length: 30 }, (_, i) => ({
        text: `後の台詞${i}`,
        startTime: i + 3,
        endTime: i + 4,
      })),
    ],
    currentSubtitle: { text: '最初の台詞', startTime: 0, endTime: 1 },
    currentTimeSec: 0,
    config: {
      ...SUBTITLE_DEFAULT_CONFIG.subtitleSidebar,
      enabled: true,
      layout: 'overlay',
      pauseVideoOnHover: false,
      autoScroll: true,
      css: {},
    },
  };
  Object.defineProperty(window, 'electronAPI', {
    value: {
      getSubtitleSidebarSnapshot: async () => snapshot,
      copySubtitleSidebarSelection: async (text) => {
        if (!('copyTestSelection' in window) || typeof window.copyTestSelection !== 'function')
          throw new Error('Missing test clipboard bridge');
        window.copyTestSelection(text);
      },
      getOverlayLayer: () => 'visible',
      sendMpvCommand: (command) => {
        commands.push(command);
      },
      setIgnoreMouseEvents: () => {},
    } satisfies Pick<
      ElectronAPI,
      | 'getSubtitleSidebarSnapshot'
      | 'copySubtitleSidebarSelection'
      | 'getOverlayLayer'
      | 'sendMpvCommand'
      | 'setIgnoreMouseEvents'
    >,
  });
  const ctx = {
    dom: resolveRendererDom(),
    state: createRendererState(),
    platform: resolvePlatformInfo(),
  };
  const modal = createSubtitleSidebarModal(ctx, {
    modalStateReader: { isAnyModalOpen: () => false },
  });
  modal.wireDomEvents();
  wireSubtitleSidebarSelection(ctx);
  await modal.openSubtitleSidebarModal();
  const list = ctx.dom.subtitleSidebarList;
  list.style.height = '180px';
  list.style.overflowY = 'auto';
  const selection = window.getSelection();
  if (!selection) throw new Error('Native selection unavailable');
  const textNode = (index: number) => {
    const node = list.children[index]?.querySelector('.subtitle-sidebar-text')?.firstChild;
    if (!node) throw new Error(`Missing cue ${index}`);
    return node;
  };
  const select = (backward = false) => {
    const start = textNode(0);
    const end = textNode(2);
    selection.setBaseAndExtent(backward ? end : start, 2, backward ? start : end, 2);
    document.dispatchEvent(new Event('selectionchange'));
    return getSubtitleSidebarSelection(list);
  };
  // This is the same competing action as the renderer's current-subtitle shortcut.
  let fallbackCopies = 0;
  document.addEventListener('keydown', (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'c') fallbackCopies += 1;
  });
  return {
    select,
    selected: () => getSubtitleSidebarSelection(list),
    buttonVisible: () => !ctx.dom.subtitleSidebarCopy.hidden,
    fallbackCopies: () => fallbackCopies,
    dragPoints: () =>
      [0, 2].map((index) => {
        const range = document.createRange();
        range.setStart(textNode(index), 2);
        range.collapse(true);
        const rect = range.getBoundingClientRect();
        return { x: Math.round(rect.x), y: Math.round(rect.y + rect.height / 2) };
      }),
    clickCopy: () => ctx.dom.subtitleSidebarCopy.click(),
    clickCue: () => {
      const before = commands.length;
      list.children[0]?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      return commands.length - before;
    },
    updatePlayback: async () => {
      list.scrollTop = 0;
      snapshot.currentTimeSec = 25;
      snapshot.currentSubtitle = { text: '後の台詞22', startTime: 25, endTime: 26 };
      await modal.refreshSubtitleSidebarSnapshot();
      return list.scrollTop;
    },
    changeSource: async () => {
      snapshot.sourceKey = 'episode-2:track-1';
      await modal.refreshSubtitleSidebarSnapshot();
      return getSubtitleSidebarSelection(list);
    },
    clear: () => selection.removeAllRanges(),
    close: () => modal.closeSubtitleSidebarModal(),
  };
}
