import type { ResolvedConfig } from '../../types/config';
import type { OverlayNotificationPayload } from '../../types/notification';

export function withConfiguredOverlayNotificationPosition(
  payload: OverlayNotificationPayload,
  config: Pick<ResolvedConfig, 'notifications'>,
): OverlayNotificationPayload {
  return {
    ...payload,
    position: payload.position ?? config.notifications.overlayPosition,
  };
}
