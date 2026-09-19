export const MEDIA_KINDS = ['anime', 'youtube'] as const;

export type MediaKind = (typeof MEDIA_KINDS)[number];
