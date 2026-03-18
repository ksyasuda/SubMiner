export const YOMITAN_POPUP_IFRAME_SELECTOR = 'iframe.yomitan-popup, iframe[id^="yomitan-popup"]';
export const YOMITAN_POPUP_SHOWN_EVENT = 'yomitan-popup-shown';
export const YOMITAN_POPUP_HIDDEN_EVENT = 'yomitan-popup-hidden';
export const YOMITAN_POPUP_MOUSE_ENTER_EVENT = 'yomitan-popup-mouse-enter';
export const YOMITAN_POPUP_MOUSE_LEAVE_EVENT = 'yomitan-popup-mouse-leave';
export const YOMITAN_POPUP_COMMAND_EVENT = 'subminer-yomitan-popup-command';
export const YOMITAN_LOOKUP_EVENT = 'subminer-yomitan-lookup';

export function registerYomitanLookupListener(
  target: EventTarget = window,
  listener: () => void,
): () => void {
  const wrapped = (): void => {
    listener();
  };
  target.addEventListener(YOMITAN_LOOKUP_EVENT, wrapped);
  return () => {
    target.removeEventListener(YOMITAN_LOOKUP_EVENT, wrapped);
  };
}

export function isYomitanPopupIframe(element: Element | null): boolean {
  if (!element) return false;
  if (element.tagName.toUpperCase() !== 'IFRAME') return false;

  const hasModernPopupClass = element.classList?.contains('yomitan-popup') ?? false;
  const hasLegacyPopupId = (element.id ?? '').startsWith('yomitan-popup');
  return hasModernPopupClass || hasLegacyPopupId;
}

export function hasYomitanPopupIframe(root: ParentNode = document): boolean {
  return root.querySelector(YOMITAN_POPUP_IFRAME_SELECTOR) !== null;
}

export function isYomitanPopupVisible(root: ParentNode = document): boolean {
  const popupIframes = root.querySelectorAll<HTMLIFrameElement>(YOMITAN_POPUP_IFRAME_SELECTOR);
  for (const iframe of popupIframes) {
    const rect = iframe.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) {
      continue;
    }
    const styles = window.getComputedStyle(iframe);
    if (styles.visibility === 'hidden' || styles.display === 'none' || styles.opacity === '0') {
      continue;
    }
    return true;
  }
  return false;
}
