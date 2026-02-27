export const YOMITAN_POPUP_IFRAME_SELECTOR = 'iframe.yomitan-popup, iframe[id^="yomitan-popup"]';
export const YOMITAN_POPUP_SHOWN_EVENT = 'yomitan-popup-shown';
export const YOMITAN_POPUP_HIDDEN_EVENT = 'yomitan-popup-hidden';

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
