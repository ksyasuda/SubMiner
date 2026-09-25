import { DEFAULT_CONFIG } from '../../definitions';
import type { ResolveContext } from '../context';
import { asBoolean, asString } from '../shared';
import { applyModernValue, asPositiveNumber } from './modern-value';

export function applyAnkiBaseResolution(
  context: ResolveContext,
  ankiConnect: Record<string, unknown>,
): void {
  applyModernValue(
    context,
    ankiConnect,
    'enabled',
    'ankiConnect.enabled',
    asBoolean,
    DEFAULT_CONFIG.ankiConnect.enabled,
    (value) => {
      context.resolved.ankiConnect.enabled = value;
    },
    'Expected boolean.',
  );
  applyModernValue(
    context,
    ankiConnect,
    'url',
    'ankiConnect.url',
    asString,
    DEFAULT_CONFIG.ankiConnect.url,
    (value) => {
      context.resolved.ankiConnect.url = value;
    },
    'Expected string.',
  );
  applyModernValue(
    context,
    ankiConnect,
    'pollingRate',
    'ankiConnect.pollingRate',
    asPositiveNumber,
    DEFAULT_CONFIG.ankiConnect.pollingRate,
    (value) => {
      context.resolved.ankiConnect.pollingRate = value;
    },
    'Expected positive number.',
  );
  applyModernValue(
    context,
    ankiConnect,
    'deck',
    'ankiConnect.deck',
    asString,
    DEFAULT_CONFIG.ankiConnect.deck,
    (value) => {
      context.resolved.ankiConnect.deck = value;
    },
    'Expected string.',
  );
}
