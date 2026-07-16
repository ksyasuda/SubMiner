import { isNotificationType, type NotificationType } from '../../../types/notification';

export function asNotificationType(value: unknown): NotificationType | undefined {
  return isNotificationType(value) ? value : undefined;
}

export function hasOwn(obj: Record<string, unknown>, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(obj, key);
}
