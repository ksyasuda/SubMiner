export const DICTIONARY_EXTERNAL_LINK_CHANNEL = 'dictionary:open-external';

export function parseDictionaryExternalUrl(value: unknown): string {
  if (typeof value !== 'string') throw new Error('Expected a link URL');
  const url = new URL(value);
  if (!['http:', 'https:'].includes(url.protocol) || url.username || url.password) {
    throw new Error('Dictionary links must use HTTP or HTTPS without credentials');
  }
  return url.href;
}
