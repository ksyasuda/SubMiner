export const YOMITAN_POPUP_IFRAME_SELECTOR = 'iframe.yomitan-popup, iframe[id^="yomitan-popup"]';
export const YOMITAN_POPUP_HOST_SELECTOR = '[data-subminer-yomitan-popup-host="true"]';
export const YOMITAN_POPUP_VISIBLE_HOST_SELECTOR =
  '[data-subminer-yomitan-popup-host="true"][data-subminer-yomitan-popup-visible="true"]';
const YOMITAN_POPUP_VISIBLE_ATTRIBUTE = 'data-subminer-yomitan-popup-visible';
export const YOMITAN_POPUP_SHOWN_EVENT = 'yomitan-popup-shown';
export const YOMITAN_POPUP_HIDDEN_EVENT = 'yomitan-popup-hidden';
export const YOMITAN_POPUP_MOUSE_ENTER_EVENT = 'yomitan-popup-mouse-enter';
export const YOMITAN_POPUP_MOUSE_LEAVE_EVENT = 'yomitan-popup-mouse-leave';
export const YOMITAN_POPUP_COMMAND_EVENT = 'subminer-yomitan-popup-command';
export const YOMITAN_LOOKUP_EVENT = 'subminer-yomitan-lookup';
export const PRIMARY_SUB_VISIBLE_ON_YOMITAN_POPUP_CLASS = 'primary-sub-visible-on-yomitan-popup';

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

export function hasYomitanPopupIframe(root: ParentNode | null | undefined = document): boolean {
  return (
    typeof root?.querySelector === 'function' &&
    (root.querySelector(YOMITAN_POPUP_IFRAME_SELECTOR) !== null ||
      root.querySelector(YOMITAN_POPUP_HOST_SELECTOR) !== null)
  );
}

function isVisiblePopupElement(element: Element): boolean {
  const rect = element.getBoundingClientRect();
  if (rect.width <= 0 || rect.height <= 0) {
    return false;
  }

  const styles = window.getComputedStyle(element);
  if (styles.visibility === 'hidden' || styles.display === 'none' || styles.opacity === '0') {
    return false;
  }

  return true;
}

function isMarkedVisiblePopupHost(element: Element): boolean {
  return element.getAttribute(YOMITAN_POPUP_VISIBLE_ATTRIBUTE) === 'true';
}

function queryPopupElements<T extends Element>(
  root: ParentNode | null | undefined,
  selector: string,
): T[] {
  if (typeof root?.querySelectorAll === 'function') {
    return Array.from(root.querySelectorAll<T>(selector));
  }
  if (typeof root?.querySelector === 'function') {
    const first = root.querySelector(selector) as T | null;
    return first ? [first] : [];
  }
  return [];
}

export function isYomitanPopupVisible(root: ParentNode | null | undefined = document): boolean {
  const visiblePopupHosts = queryPopupElements<HTMLElement>(
    root,
    YOMITAN_POPUP_VISIBLE_HOST_SELECTOR,
  );
  if (visiblePopupHosts.length > 0) {
    return true;
  }

  const popupIframes = queryPopupElements<HTMLIFrameElement>(root, YOMITAN_POPUP_IFRAME_SELECTOR);
  for (const iframe of popupIframes) {
    if (isVisiblePopupElement(iframe)) {
      return true;
    }
  }

  const popupHosts = queryPopupElements<HTMLElement>(root, YOMITAN_POPUP_HOST_SELECTOR);
  for (const host of popupHosts) {
    if (isMarkedVisiblePopupHost(host)) {
      return true;
    }
  }
  return false;
}
