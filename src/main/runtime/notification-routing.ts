import type { NotificationType } from '../../types/notification';

export function shouldShowOsd(type: NotificationType): boolean {
  return type === 'osd' || type === 'osd-system';
}

export function shouldShowOverlay(type: NotificationType): boolean {
  return type === 'overlay' || type === 'both';
}

export function shouldShowDesktop(type: NotificationType): boolean {
  return type === 'system' || type === 'both' || type === 'osd-system';
}

export function resolveOverlayReadinessNotificationType(
  type: NotificationType,
  _overlayReady: boolean,
): NotificationType {
  return type;
}
