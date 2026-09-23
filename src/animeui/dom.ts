/** Small helpers shared by the anime browser's panels. */

export function el<T extends HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (!node) throw new Error(`Missing element #${id}`);
  return node as T;
}

export function describe(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  // Electron wraps handler errors; keep only the useful tail.
  return message.replace(/^Error invoking remote method '[^']+':\s*/, '').replace(/^Error:\s*/, '');
}
