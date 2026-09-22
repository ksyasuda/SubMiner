import type {
  CompiledSessionBinding,
  SessionBindingWarning,
  SessionKeySpec,
} from '../types/session-bindings';

export interface SessionKeyReservation {
  key: SessionKeySpec;
  path: string;
}

export function getSessionSequencePrefix(key: SessionKeySpec): SessionKeySpec | null {
  const match = /^(Key[A-Z])-Key[A-Z]$/.exec(key.code);
  return match?.[1] ? { code: match[1], modifiers: key.modifiers } : null;
}

function signature(key: SessionKeySpec): string {
  return [...key.modifiers, key.code].join('+');
}

export function resolveSessionSequenceConflicts(
  bindings: CompiledSessionBinding[],
  reservations: SessionKeyReservation[] = [],
): { bindings: CompiledSessionBinding[]; warnings: SessionBindingWarning[] } {
  const singles = new Map<string, string[]>();
  for (const { key, path } of [
    ...bindings.map((binding) => ({ key: binding.key, path: binding.sourcePath })),
    ...reservations,
  ]) {
    if (getSessionSequencePrefix(key)) continue;
    const id = signature(key);
    singles.set(id, [...(singles.get(id) ?? []), path]);
  }
  const warnings: SessionBindingWarning[] = [];
  const effective = bindings.filter((binding) => {
    const prefix = getSessionSequencePrefix(binding.key);
    if (!prefix) return true;
    const conflicts = singles.get(signature(prefix));
    if (!conflicts?.length) return true;
    const paths = [...new Set(conflicts)];
    warnings.push({
      kind: 'conflict',
      path: binding.sourcePath,
      value: binding.originalKey,
      conflictingPaths: paths,
      message: `Disabled sequence "${binding.originalKey}" (${binding.sourcePath}): its first key is reserved by ${paths.join(', ')}. Single-key bindings take priority; remap the sequence or its conflicting binding.`,
    });
    return false;
  });
  return { bindings: effective, warnings };
}

// Imported mpv keys preserve case: g and G are different strokes.
export function reserveMpvSequencePrefixes(keys: string[]): SessionKeyReservation[] {
  return keys.flatMap((value) => {
    const parts = value.split('+');
    const letter = parts.pop();
    if (!letter || !/^[a-z]$/i.test(letter)) return [];
    const modifiers: SessionKeySpec['modifiers'] = [];
    if (parts.includes('ctrl')) modifiers.push('ctrl');
    if (parts.includes('alt')) modifiers.push('alt');
    if (parts.includes('shift') || /^[A-Z]$/.test(letter)) modifiers.push('shift');
    if (parts.includes('meta')) modifiers.push('meta');
    return [
      {
        key: { code: `Key${letter.toUpperCase()}`, modifiers },
        path: `mpv input binding "${value}"`,
      },
    ];
  });
}
