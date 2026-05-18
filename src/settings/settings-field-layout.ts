import type { ConfigSettingsField } from '../types/settings';

export interface FieldTitleBadge {
  className: string;
  text: string;
}

export function getFieldTitleBadges(field: ConfigSettingsField): FieldTitleBadge[] {
  return [
    {
      className: `restart-chip ${field.restartBehavior}`,
      text: field.restartBehavior === 'hot-reload' ? 'Live' : 'Restart',
    },
  ];
}
