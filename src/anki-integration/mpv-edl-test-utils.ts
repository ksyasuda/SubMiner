export function toMpvEdlValue(value: string): string {
  return `%${Buffer.byteLength(value, 'utf8')}%${value}`;
}
