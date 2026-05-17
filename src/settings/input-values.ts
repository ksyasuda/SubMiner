export type OptionalNumberInputParseResult =
  | {
      ok: true;
      value: number | undefined;
    }
  | {
      ok: false;
    };

export function parseOptionalNumberInputValue(value: string): OptionalNumberInputParseResult {
  const raw = value.trim();
  if (raw.length === 0) {
    return { ok: true, value: undefined };
  }
  const next = Number(raw);
  if (!Number.isFinite(next)) {
    return { ok: false };
  }
  return { ok: true, value: next };
}
