/*
 * SubMiner - All-in-one sentence mining overlay
 * Copyright (C) 2024 sudacode
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

interface MergedToken {
  surface: string;
  reading: string;
  headword: string;
  startPos: number;
  endPos: number;
  partOfSpeech: string;
  isMerged: boolean;
}

interface SubtitleData {
  text: string;
  tokens: MergedToken[] | null;
}

interface Keybinding {
  key: string;
  command: string[] | null;
}

interface SubtitlePosition {
  yPercent: number;
}

const subtitleRoot = document.getElementById('subtitleRoot')!;
const subtitleContainer = document.getElementById('subtitleContainer')!;
const overlay = document.getElementById('overlay')!;

let isOverSubtitle = false;
let isDragging = false;
let dragStartY = 0;
let startYPercent = 0;

function normalizeSubtitle(text: string): string {
  if (!text) return '';

  let normalized = text
    .replace(/\\N/g, '\n')
    .replace(/\\n/g, '\n');

  normalized = normalized.replace(/\{[^}]*\}/g, '');

  return normalized.trim();
}

function renderWithTokens(tokens: MergedToken[]): void {
  const fragment = document.createDocumentFragment();

  for (const token of tokens) {
    const surface = token.surface;

    if (surface.includes('\n')) {
      const parts = surface.split('\n');
      for (let i = 0; i < parts.length; i++) {
        if (parts[i]) {
          const span = document.createElement('span');
          span.className = 'word';
          span.textContent = parts[i];
          if (token.reading) {
            span.dataset.reading = token.reading;
          }
          if (token.headword) {
            span.dataset.headword = token.headword;
          }
          fragment.appendChild(span);
        }
        if (i < parts.length - 1) {
          fragment.appendChild(document.createElement('br'));
        }
      }
    } else {
      const span = document.createElement('span');
      span.className = 'word';
      span.textContent = surface;
      if (token.reading) {
        span.dataset.reading = token.reading;
      }
      if (token.headword) {
        span.dataset.headword = token.headword;
      }
      fragment.appendChild(span);
    }
  }

  subtitleRoot.appendChild(fragment);
}

function renderCharacterLevel(text: string): void {
  const fragment = document.createDocumentFragment();

  for (const char of text) {
    if (char === '\n') {
      fragment.appendChild(document.createElement('br'));
    } else {
      const span = document.createElement('span');
      span.className = 'c';
      span.textContent = char;
      fragment.appendChild(span);
    }
  }

  subtitleRoot.appendChild(fragment);
}

function renderSubtitle(data: SubtitleData | string): void {
  subtitleRoot.innerHTML = '';

  let text: string;
  let tokens: MergedToken[] | null;

  if (typeof data === 'string') {
    text = data;
    tokens = null;
  } else if (data && typeof data === 'object') {
    text = data.text;
    tokens = data.tokens;
  } else {
    return;
  }

  if (!text) {
    return;
  }

  const normalized = normalizeSubtitle(text);

  if (tokens && tokens.length > 0) {
    renderWithTokens(tokens);
  } else {
    renderCharacterLevel(normalized);
  }
}

function handleMouseEnter(): void {
  isOverSubtitle = true;
  overlay.classList.add('interactive');
}

function handleMouseLeave(): void {
  isOverSubtitle = false;
  const yomitanPopup = document.querySelector('iframe[id^="yomitan-popup"]');
  if (!yomitanPopup) {
    overlay.classList.remove('interactive');
  }
}

function getCurrentYPercent(): number {
  const marginBottom = parseFloat(subtitleContainer.style.marginBottom) || 60;
  const windowHeight = window.innerHeight;
  return (marginBottom / windowHeight) * 100;
}

function applyYPercent(yPercent: number): void {
  const clampedPercent = Math.max(2, Math.min(80, yPercent));
  const marginBottom = (clampedPercent / 100) * window.innerHeight;

  subtitleContainer.style.position = '';
  subtitleContainer.style.left = '';
  subtitleContainer.style.top = '';
  subtitleContainer.style.right = '';
  subtitleContainer.style.transform = '';

  subtitleContainer.style.marginBottom = `${marginBottom}px`;
}

function applyStoredSubtitlePosition(position: SubtitlePosition | null, source: string): void {
  if (position && position.yPercent !== undefined) {
    applyYPercent(position.yPercent);
    console.log('Applied subtitle position from', source, ':', position.yPercent, '%');
  } else {
    const defaultMarginBottom = 60;
    const defaultYPercent = (defaultMarginBottom / window.innerHeight) * 100;
    applyYPercent(defaultYPercent);
    console.log('Applied default subtitle position from', source);
  }
}

function applySubtitleFontSize(fontSize: number): void {
  const clampedSize = Math.max(10, Math.min(96, fontSize));
  document.documentElement.style.setProperty(
    '--subtitle-font-size',
    `${clampedSize}px`,
  );
}

function setupDragging(): void {
  subtitleContainer.addEventListener('mousedown', (e: MouseEvent) => {
    if (e.button === 2) {
      e.preventDefault();
      isDragging = true;
      dragStartY = e.clientY;
      startYPercent = getCurrentYPercent();
      subtitleContainer.style.cursor = 'grabbing';
    }
  });

  document.addEventListener('mousemove', (e: MouseEvent) => {
    if (!isDragging) return;

    const deltaY = dragStartY - e.clientY;
    const deltaPercent = (deltaY / window.innerHeight) * 100;
    const newYPercent = startYPercent + deltaPercent;

    applyYPercent(newYPercent);
  });

  document.addEventListener('mouseup', (e: MouseEvent) => {
    if (isDragging && e.button === 2) {
      isDragging = false;
      subtitleContainer.style.cursor = '';

      const yPercent = getCurrentYPercent();
      window.electronAPI.saveSubtitlePosition({ yPercent });
    }
  });

  subtitleContainer.addEventListener('contextmenu', (e: Event) => {
    e.preventDefault();
  });
}

function isInteractiveTarget(target: EventTarget | null): boolean {
  if (!(target instanceof Element)) return false;
  if (subtitleContainer.contains(target)) return true;
  if (target.tagName === 'IFRAME' && target.id && target.id.startsWith('yomitan-popup')) return true;
  if (target.closest && target.closest('iframe[id^="yomitan-popup"]')) return true;
  return false;
}

function keyEventToString(e: KeyboardEvent): string {
  const parts: string[] = [];
  if (e.ctrlKey) parts.push('Ctrl');
  if (e.altKey) parts.push('Alt');
  if (e.shiftKey) parts.push('Shift');
  if (e.metaKey) parts.push('Meta');
  parts.push(e.code);
  return parts.join('+');
}

let keybindingsMap = new Map<string, string[]>();

type ChordAction =
  | { type: 'mpv'; command: string[] }
  | { type: 'electron'; action: () => void }
  | { type: 'noop' };

const CHORD_MAP = new Map<string, ChordAction>([
  ['KeyT', { type: 'electron', action: () => window.electronAPI.toggleOverlay() }],
  ['Shift+KeyS', { type: 'electron', action: () => window.electronAPI.quitApp() }],
  ['KeyO', { type: 'electron', action: () => window.electronAPI.openYomitanSettings() }],
  ['KeyR', { type: 'mpv', command: ['script-message', 'subminer-restart'] }],
  ['KeyC', { type: 'mpv', command: ['script-message', 'subminer-status'] }],
  ['KeyY', { type: 'mpv', command: ['script-message', 'subminer-menu'] }],
  ['KeyD', { type: 'electron', action: () => window.electronAPI.toggleDevTools() }],
  ['KeyS', { type: 'noop' }],
]);

let chordPending = false;
let chordTimeout: ReturnType<typeof setTimeout> | null = null;

function resetChord(): void {
  chordPending = false;
  if (chordTimeout !== null) {
    clearTimeout(chordTimeout);
    chordTimeout = null;
  }
}

async function setupMpvInputForwarding(): Promise<void> {
  const keybindings: Keybinding[] = await window.electronAPI.getKeybindings();
  keybindingsMap = new Map();
  for (const binding of keybindings) {
    if (binding.command) {
      keybindingsMap.set(binding.key, binding.command);
    }
  }

  document.addEventListener('keydown', (e: KeyboardEvent) => {
    const yomitanPopup = document.querySelector('iframe[id^="yomitan-popup"]');
    if (yomitanPopup) return;

    if (chordPending) {
      const modifierKeys = ['ShiftLeft', 'ShiftRight', 'ControlLeft', 'ControlRight',
                            'AltLeft', 'AltRight', 'MetaLeft', 'MetaRight'];
      if (modifierKeys.includes(e.code)) {
        return;
      }

      e.preventDefault();
      const secondKey = keyEventToString(e);
      const action = CHORD_MAP.get(secondKey);
      resetChord();
      if (action) {
        if (action.type === 'mpv') {
          window.electronAPI.sendMpvCommand(action.command);
        } else if (action.type === 'electron') {
          action.action();
        }
      }
      return;
    }

    if (e.code === 'KeyY' && !e.ctrlKey && !e.altKey && !e.shiftKey && !e.metaKey && !e.repeat) {
      e.preventDefault();
      chordPending = true;
      chordTimeout = setTimeout(() => {
        resetChord();
      }, 1000);
      return;
    }

    const keyString = keyEventToString(e);
    const command = keybindingsMap.get(keyString);

    if (command) {
      e.preventDefault();
      window.electronAPI.sendMpvCommand(command);
    }
  });

  document.addEventListener('mousedown', (e: MouseEvent) => {
    if (e.button === 2 && !isInteractiveTarget(e.target)) {
      e.preventDefault();
      window.electronAPI.sendMpvCommand(['cycle', 'pause']);
    }
  });

  document.addEventListener('contextmenu', (e: Event) => {
    if (!isInteractiveTarget(e.target)) {
      e.preventDefault();
    }
  });
}

function setupResizeHandler(): void {
  window.addEventListener('resize', () => {
    const currentYPercent = getCurrentYPercent();
    applyYPercent(currentYPercent);
  });
}

async function restoreSubtitlePosition(): Promise<void> {
  const position = await window.electronAPI.getSubtitlePosition();
  applyStoredSubtitlePosition(position, 'startup');
}

async function restoreSubtitleFontSize(): Promise<void> {
  const style = await window.electronAPI.getSubtitleStyle();
  applySubtitleFontSize(style.fontSize);
  console.log('Applied subtitle font size:', style.fontSize);
}

function setupSelectionObserver(): void {
  document.addEventListener('selectionchange', () => {
    const selection = window.getSelection();
    const hasSelection = selection && selection.rangeCount > 0 && !selection.isCollapsed;

    if (hasSelection) {
      subtitleRoot.classList.add('has-selection');
    } else {
      subtitleRoot.classList.remove('has-selection');
    }
  });
}

function setupYomitanObserver(): void {
  const observer = new MutationObserver((mutations: MutationRecord[]) => {
    for (const mutation of mutations) {
      mutation.addedNodes.forEach((node) => {
        if (node.nodeType === Node.ELEMENT_NODE) {
          const element = node as Element;
          if (element.tagName === 'IFRAME' && element.id && element.id.startsWith('yomitan-popup')) {
            overlay.classList.add('interactive');
          }
        }
      });
      mutation.removedNodes.forEach((node) => {
        if (node.nodeType === Node.ELEMENT_NODE) {
          const element = node as Element;
          if (element.tagName === 'IFRAME' && element.id && element.id.startsWith('yomitan-popup')) {
            if (!isOverSubtitle) {
              overlay.classList.remove('interactive');
            }
          }
        }
      });
    }
  });

  observer.observe(document.body, {
    childList: true,
    subtree: true
  });
}

async function init(): Promise<void> {
  window.electronAPI.onSubtitle((data: SubtitleData) => {
    renderSubtitle(data);
  });

  window.electronAPI.onSubtitlePosition((position: SubtitlePosition | null) => {
    applyStoredSubtitlePosition(position, 'media-change');
  });

  const initialSubtitle = await window.electronAPI.getCurrentSubtitle();
  renderSubtitle(initialSubtitle);

  subtitleContainer.addEventListener('mouseenter', handleMouseEnter);
  subtitleContainer.addEventListener('mouseleave', handleMouseLeave);

  setupDragging();

  await setupMpvInputForwarding();

  setupResizeHandler();

  await restoreSubtitlePosition();
  await restoreSubtitleFontSize();

  setupYomitanObserver();

  setupSelectionObserver();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
