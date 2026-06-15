import type { ConfigSettingsField } from '../types/settings';
import { i18n } from '../i18n/index.js';

export interface FieldTitleBadge {
  className: string;
  text: string;
}

export function getFieldTitleBadges(field: ConfigSettingsField): FieldTitleBadge[] {
  return [
    {
      className: `restart-chip ${field.restartBehavior}`,
      text: field.restartBehavior === 'hot-reload' ? i18n.t('settingsBadge.live') : i18n.t('settingsBadge.restart'),
    },
  ];
}
