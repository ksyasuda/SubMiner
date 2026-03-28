/*
 * SubMiner - All-in-one sentence mining overlay
 * Copyright (C) 2024 sudacode
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

import { AnkiConnectConfig } from './types/anki';
import {
  RuntimeOptionApplyResult,
  RuntimeOptionId,
  RuntimeOptionState,
  RuntimeOptionValue,
} from './types/runtime-options';
import { SubtitleStyleConfig } from './types/subtitle';
import { RUNTIME_OPTION_REGISTRY, RuntimeOptionRegistryEntry } from './config';

type RuntimeOverrides = Record<string, unknown>;

function deepClone<T>(value: T): T {
  return structuredClone(value);
}

function getPathValue(source: Record<string, unknown>, path: string): unknown {
  const parts = path.split('.');
  let current: unknown = source;
  for (const part of parts) {
    if (!current || typeof current !== 'object' || Array.isArray(current)) {
      return undefined;
    }
    current = (current as Record<string, unknown>)[part];
  }
  return current;
}

function setPathValue(target: Record<string, unknown>, path: string, value: unknown): void {
  const parts = path.split('.');
  let current = target;
  for (let i = 0; i < parts.length; i += 1) {
    const part = parts[i]!;
    const isLeaf = i === parts.length - 1;
    if (isLeaf) {
      current[part] = value;
      return;
    }

    const next = current[part];
    if (!next || typeof next !== 'object' || Array.isArray(next)) {
      current[part] = {};
    }
    current = current[part] as Record<string, unknown>;
  }
}

function allowedValues(definition: RuntimeOptionRegistryEntry): RuntimeOptionValue[] {
  return [...definition.allowedValues];
}

function isAllowedValue(
  definition: RuntimeOptionRegistryEntry,
  value: RuntimeOptionValue,
): boolean {
  if (definition.valueType === 'boolean') {
    return typeof value === 'boolean';
  }
  return typeof value === 'string' && definition.allowedValues.includes(value);
}

export class RuntimeOptionsManager {
  private readonly getAnkiConfig: () => AnkiConnectConfig;
  private readonly getSubtitleStyleConfig?: () => SubtitleStyleConfig;
  private readonly applyAnkiPatch: (patch: Partial<AnkiConnectConfig>) => void;
  private readonly onOptionsChanged: (options: RuntimeOptionState[]) => void;
  private runtimeOverrides: RuntimeOverrides = {};
  private readonly definitions = new Map<RuntimeOptionId, RuntimeOptionRegistryEntry>();

  constructor(
    getAnkiConfig: () => AnkiConnectConfig,
    callbacks: {
      applyAnkiPatch: (patch: Partial<AnkiConnectConfig>) => void;
      onOptionsChanged: (options: RuntimeOptionState[]) => void;
      getSubtitleStyleConfig?: () => SubtitleStyleConfig;
    },
  ) {
    this.getAnkiConfig = getAnkiConfig;
    this.getSubtitleStyleConfig = callbacks.getSubtitleStyleConfig;
    this.applyAnkiPatch = callbacks.applyAnkiPatch;
    this.onOptionsChanged = callbacks.onOptionsChanged;
    for (const definition of RUNTIME_OPTION_REGISTRY) {
      this.definitions.set(definition.id, definition);
    }
  }

  private getEffectiveValue(definition: RuntimeOptionRegistryEntry): RuntimeOptionValue {
    const override = getPathValue(this.runtimeOverrides, definition.path);
    if (override !== undefined) return override as RuntimeOptionValue;

    const source = {
      ankiConnect: this.getAnkiConfig(),
      subtitleStyle: this.getSubtitleStyleConfig?.(),
    } as Record<string, unknown>;

    const raw = getPathValue(source, definition.path);
    if (raw === undefined || raw === null) {
      return definition.defaultValue;
    }
    return raw as RuntimeOptionValue;
  }

  listOptions(): RuntimeOptionState[] {
    const options: RuntimeOptionState[] = [];
    for (const definition of RUNTIME_OPTION_REGISTRY) {
      options.push({
        id: definition.id,
        label: definition.label,
        scope: definition.scope,
        valueType: definition.valueType,
        value: this.getEffectiveValue(definition),
        allowedValues: allowedValues(definition),
        requiresRestart: definition.requiresRestart,
      });
    }
    return options;
  }

  getOptionValue(id: RuntimeOptionId): RuntimeOptionValue | undefined {
    const definition = this.definitions.get(id);
    if (!definition) return undefined;
    return this.getEffectiveValue(definition);
  }

  setOptionValue(id: RuntimeOptionId, value: RuntimeOptionValue): RuntimeOptionApplyResult {
    const definition = this.definitions.get(id);
    if (!definition) {
      return { ok: false, error: `Unknown runtime option: ${id}` };
    }

    if (!isAllowedValue(definition, value)) {
      return {
        ok: false,
        error: `Invalid value for ${id}: ${String(value)}`,
      };
    }

    const next = deepClone(this.runtimeOverrides);
    setPathValue(next, definition.path, value);
    this.runtimeOverrides = next;

    if (definition.scope === 'ankiConnect') {
      const ankiPatch = definition.toAnkiPatch(value);
      this.applyAnkiPatch(ankiPatch);
    }

    const option = this.listOptions().find((item) => item.id === id);
    if (!option) {
      return { ok: false, error: `Failed to apply option: ${id}` };
    }

    const osdMessage = `Runtime option: ${definition.label} -> ${definition.formatValueForOsd(option.value)}`;
    this.onOptionsChanged(this.listOptions());
    return {
      ok: true,
      option,
      osdMessage,
      requiresRestart: definition.requiresRestart,
    };
  }

  cycleOption(id: RuntimeOptionId, direction: 1 | -1): RuntimeOptionApplyResult {
    const definition = this.definitions.get(id);
    if (!definition) {
      return { ok: false, error: `Unknown runtime option: ${id}` };
    }

    const values = allowedValues(definition);
    if (values.length === 0) {
      return { ok: false, error: `Option ${id} has no allowed values` };
    }

    const currentValue = this.getEffectiveValue(definition);
    const currentIndex = values.findIndex((value) => value === currentValue);
    const safeIndex = currentIndex >= 0 ? currentIndex : 0;
    const nextIndex =
      direction === 1
        ? (safeIndex + 1) % values.length
        : (safeIndex - 1 + values.length) % values.length;
    return this.setOptionValue(id, values[nextIndex]!);
  }

  getEffectiveAnkiConnectConfig(baseConfig?: AnkiConnectConfig): AnkiConnectConfig {
    const source = baseConfig ?? this.getAnkiConfig();
    const effective: AnkiConnectConfig = deepClone(source);

    for (const definition of RUNTIME_OPTION_REGISTRY) {
      const override = getPathValue(this.runtimeOverrides, definition.path);
      if (override === undefined) continue;
      if (!definition.path.startsWith('ankiConnect.')) continue;

      const subPath = definition.path.replace(/^ankiConnect\./, '');
      setPathValue(effective as unknown as Record<string, unknown>, subPath, override);
    }

    return effective;
  }
}
