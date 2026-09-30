import { ipcRenderer } from 'electron';

// Sandboxed preloads can only require `electron`, so these mirror IPC_CHANNELS in
// src/shared/ipc/contracts.ts (youtubeBrowserOpenVideo / youtubeBrowserToast).
const OPEN_VIDEO_CHANNEL = 'youtube-browser:open-video';
const TOAST_CHANNEL = 'youtube-browser:toast';

const MOUSE_BUTTON_MIDDLE = 1;
const MOUSE_BUTTON_BACK = 3;
const MOUSE_BUTTON_FORWARD = 4;

function isYoutubeVideoHref(href: string): boolean {
  let url: URL;
  try {
    url = new URL(href);
  } catch {
    return false;
  }
  const host = url.hostname.toLowerCase();
  if (host !== 'youtube.com' && !host.endsWith('.youtube.com')) return false;
  return (
    (url.pathname === '/watch' && url.searchParams.has('v')) ||
    /^\/shorts\/[^/]+/.test(url.pathname)
  );
}

function findVideoLink(event: MouseEvent): HTMLAnchorElement | null {
  for (const node of event.composedPath()) {
    if (node instanceof HTMLAnchorElement && node.href && isYoutubeVideoHref(node.href)) {
      return node;
    }
  }
  return null;
}

// Capture phase runs before YouTube's own click handlers, so the page never starts its player.
// Plain click plays now; middle click or a modifier-click queues behind the current video.
function interceptVideoClick(event: MouseEvent): void {
  const isPrimary = event.type === 'click' && event.button === 0;
  const isMiddle = event.type === 'auxclick' && event.button === MOUSE_BUTTON_MIDDLE;
  if (!isPrimary && !isMiddle) return;
  const link = findVideoLink(event);
  if (!link) return;
  event.preventDefault();
  event.stopImmediatePropagation();
  const queue = isMiddle || event.shiftKey || event.ctrlKey || event.metaKey;
  ipcRenderer.send(OPEN_VIDEO_CHANNEL, { action: queue ? 'queue' : 'play', url: link.href });
  // Immediate feedback: starting mpv can take a few seconds before the result toast arrives.
  showToast(queue ? 'Queueing in mpv…' : 'Starting mpv…', true);
}

window.addEventListener('click', interceptVideoClick, true);
window.addEventListener('auxclick', interceptVideoClick, true);

// Chromium only maps mouse back/forward buttons to history inside a real browser.
window.addEventListener('mouseup', (event) => {
  if (event.button === MOUSE_BUTTON_BACK) history.back();
  if (event.button === MOUSE_BUTTON_FORWARD) history.forward();
});

let toastTimer: ReturnType<typeof setTimeout> | null = null;

function showToast(message: string, ok: boolean): void {
  let toast = document.getElementById('subminer-youtube-toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'subminer-youtube-toast';
    Object.assign(toast.style, {
      position: 'fixed',
      left: '50%',
      bottom: '32px',
      transform: 'translateX(-50%)',
      zIndex: '2147483647',
      padding: '10px 18px',
      borderRadius: '8px',
      font: '500 14px/1.4 Roboto, Arial, sans-serif',
      color: '#fff',
      boxShadow: '0 4px 16px rgba(0, 0, 0, 0.4)',
      pointerEvents: 'none',
      transition: 'opacity 150ms ease',
    });
    document.body.appendChild(toast);
  }
  toast.textContent = message;
  toast.style.background = ok ? '#1f1f1f' : '#8c1d18';
  toast.style.opacity = '1';
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    if (toast) toast.style.opacity = '0';
  }, 2200);
}

ipcRenderer.on(TOAST_CHANNEL, (_event, payload: unknown) => {
  if (typeof payload !== 'object' || payload === null) return;
  const { message, ok } = payload as { message?: unknown; ok?: unknown };
  if (typeof message === 'string') showToast(message, ok !== false);
});
