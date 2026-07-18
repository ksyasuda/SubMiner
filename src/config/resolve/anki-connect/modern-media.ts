import { DEFAULT_CONFIG } from '../../definitions';
import type { ResolveContext } from '../context';
import { asBoolean } from '../shared';
import {
  applyModernValue,
  asIntegerInRange,
  asNonNegativeInteger,
  asNonNegativeNumber,
  asPositiveNumber,
} from './modern-value';

export function applyModernMediaResolution(
  context: ResolveContext,
  media: Record<string, unknown>,
): void {
  for (const key of [
    'generateAudio',
    'generateImage',
    'syncAnimatedImageToWordAudio',
    'normalizeAudio',
    'mirrorMpvVolume',
  ] as const) {
    applyModernValue(
      context,
      media,
      key,
      `ankiConnect.media.${key}`,
      asBoolean,
      DEFAULT_CONFIG.ankiConnect.media[key],
      (value) => {
        context.resolved.ankiConnect.media[key] = value;
      },
      'Expected boolean.',
    );
  }

  applyModernValue(
    context,
    media,
    'imageType',
    'ankiConnect.media.imageType',
    (value) => (value === 'static' || value === 'avif' ? value : undefined),
    DEFAULT_CONFIG.ankiConnect.media.imageType,
    (value) => {
      context.resolved.ankiConnect.media.imageType = value;
    },
    "Expected 'static' or 'avif'.",
  );
  applyModernValue(
    context,
    media,
    'imageFormat',
    'ankiConnect.media.imageFormat',
    (value) => (value === 'jpg' || value === 'png' || value === 'webp' ? value : undefined),
    DEFAULT_CONFIG.ankiConnect.media.imageFormat,
    (value) => {
      context.resolved.ankiConnect.media.imageFormat = value;
    },
    "Expected 'jpg', 'png', or 'webp'.",
  );
  applyModernValue(
    context,
    media,
    'imageQuality',
    'ankiConnect.media.imageQuality',
    (value) => asIntegerInRange(value, 1, 100),
    DEFAULT_CONFIG.ankiConnect.media.imageQuality,
    (value) => {
      context.resolved.ankiConnect.media.imageQuality = value;
    },
    'Expected integer between 1 and 100.',
  );

  for (const key of [
    'imageMaxWidth',
    'imageMaxHeight',
    'animatedMaxWidth',
    'animatedMaxHeight',
  ] as const) {
    applyModernValue(
      context,
      media,
      key,
      `ankiConnect.media.${key}`,
      asNonNegativeInteger,
      DEFAULT_CONFIG.ankiConnect.media[key] ?? 0,
      (value) => {
        context.resolved.ankiConnect.media[key] = value;
      },
      'Expected non-negative integer.',
    );
  }

  applyModernValue(
    context,
    media,
    'animatedFps',
    'ankiConnect.media.animatedFps',
    (value) => asIntegerInRange(value, 1, 60),
    DEFAULT_CONFIG.ankiConnect.media.animatedFps,
    (value) => {
      context.resolved.ankiConnect.media.animatedFps = value;
    },
    'Expected integer between 1 and 60.',
  );
  applyModernValue(
    context,
    media,
    'animatedCrf',
    'ankiConnect.media.animatedCrf',
    (value) => asIntegerInRange(value, 0, 63),
    DEFAULT_CONFIG.ankiConnect.media.animatedCrf,
    (value) => {
      context.resolved.ankiConnect.media.animatedCrf = value;
    },
    'Expected integer between 0 and 63.',
  );
  applyModernValue(
    context,
    media,
    'audioPadding',
    'ankiConnect.media.audioPadding',
    asNonNegativeNumber,
    DEFAULT_CONFIG.ankiConnect.media.audioPadding,
    (value) => {
      context.resolved.ankiConnect.media.audioPadding = value;
    },
    'Expected non-negative number.',
  );

  for (const key of ['fallbackDuration', 'maxMediaDuration'] as const) {
    applyModernValue(
      context,
      media,
      key,
      `ankiConnect.media.${key}`,
      asPositiveNumber,
      DEFAULT_CONFIG.ankiConnect.media[key],
      (value) => {
        context.resolved.ankiConnect.media[key] = value;
      },
      'Expected positive number.',
    );
  }
}
