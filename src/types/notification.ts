export const SETTINGS_NOTIFICATION_TYPE_VALUES = ['overlay', 'system', 'both', 'none'] as const;

export const NOTIFICATION_TYPE_VALUES = [
  ...SETTINGS_NOTIFICATION_TYPE_VALUES,
  'osd',
  'osd-system',
] as const;

export const OVERLAY_NOTIFICATION_POSITION_VALUES = ['top-left', 'top', 'top-right'] as const;

export type SettingsNotificationType = (typeof SETTINGS_NOTIFICATION_TYPE_VALUES)[number];
export type NotificationType = (typeof NOTIFICATION_TYPE_VALUES)[number];
export type OverlayNotificationPosition = (typeof OVERLAY_NOTIFICATION_POSITION_VALUES)[number];

export type OverlayNotificationVariant = 'info' | 'success' | 'warning' | 'error' | 'progress';

export const OPEN_ANKI_CARD_ACTION_ID = 'open-anki-card';

export interface OverlayNotificationAction {
  id: string;
  label: string;
  noteId?: number;
  /** Leaves the notification on screen after the action fires (default: dismiss). */
  keepOpen?: boolean;
}

export interface OverlayNotificationPayload {
  id?: string;
  historyId?: string;
  title: string;
  body?: string;
  image?: string;
  variant?: OverlayNotificationVariant;
  position?: OverlayNotificationPosition;
  persistent?: boolean;
  timeoutMs?: number;
  actions?: OverlayNotificationAction[];
}

export interface OverlayNotificationDismissPayload {
  id: string;
  dismiss: true;
}

export type OverlayNotificationEventPayload =
  | OverlayNotificationPayload
  | OverlayNotificationDismissPayload;

export function isNotificationType(value: unknown): value is NotificationType {
  return typeof value === 'string' && NOTIFICATION_TYPE_VALUES.includes(value as NotificationType);
}

export function isOverlayNotificationPosition(
  value: unknown,
): value is OverlayNotificationPosition {
  return (
    typeof value === 'string' &&
    OVERLAY_NOTIFICATION_POSITION_VALUES.includes(value as OverlayNotificationPosition)
  );
}
