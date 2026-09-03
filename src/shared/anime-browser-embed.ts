export const ANIME_BROWSER_CLOSE_MESSAGE = 'subminer:anime-browser-close';

const ANIME_BROWSER_KEYDOWN_MESSAGE = 'subminer:anime-browser-keydown';

export interface AnimeBrowserKeydownMessage {
  type: typeof ANIME_BROWSER_KEYDOWN_MESSAGE;
  bindingKey: string;
  repeat: boolean;
}

export function createAnimeBrowserKeydownMessage(
  event: Pick<KeyboardEvent, 'altKey' | 'code' | 'ctrlKey' | 'metaKey' | 'repeat' | 'shiftKey'>,
): AnimeBrowserKeydownMessage {
  const parts: string[] = [];
  if (event.ctrlKey) parts.push('Ctrl');
  if (event.altKey) parts.push('Alt');
  if (event.shiftKey) parts.push('Shift');
  if (event.metaKey) parts.push('Meta');
  parts.push(event.code);
  return {
    type: ANIME_BROWSER_KEYDOWN_MESSAGE,
    bindingKey: parts.join('+'),
    repeat: event.repeat,
  };
}

export function isAnimeBrowserKeydownMessage(value: unknown): value is AnimeBrowserKeydownMessage {
  if (!value || typeof value !== 'object') return false;
  const message = value as Partial<AnimeBrowserKeydownMessage>;
  return (
    message.type === ANIME_BROWSER_KEYDOWN_MESSAGE &&
    typeof message.bindingKey === 'string' &&
    message.bindingKey.length > 0 &&
    typeof message.repeat === 'boolean'
  );
}
