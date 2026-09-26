import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';
import { parseHachidoriManagementUrl } from '../../../shared/hachidori-sharing';

// Upload from the main process: extension blob URLs cannot cross the sharing link,
// and the Docker management API deliberately rejects browser cross-origin writes.
export async function uploadHachidoriDictionary(
  zipPath: string,
  managementUrl: string,
): Promise<void> {
  const origin = parseHachidoriManagementUrl(managementUrl);
  if (!origin) {
    throw new Error(
      'Set hachidori.externalHostManagementUrl to the linked Docker host management URL to sync character dictionaries.',
    );
  }
  const url = new URL('/import', origin);
  url.searchParams.set('name', path.basename(zipPath));
  url.searchParams.set('replace', 'true');
  const bytes = await readFile(zipPath);
  const signal = AbortSignal.timeout(300_000);
  for (;;) {
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/zip' },
      body: bytes,
      signal,
      redirect: 'error',
    });
    if (response.status === 409) {
      await response.body?.cancel();
      await delay(500, undefined, { signal });
      continue;
    }
    const result: unknown = await response.json();
    if (
      response.ok &&
      typeof result === 'object' &&
      result !== null &&
      'ok' in result &&
      result.ok === true &&
      'report' in result &&
      typeof result.report === 'object' &&
      result.report !== null &&
      'success' in result.report &&
      result.report.success === true
    )
      return;
    const detail =
      typeof result === 'object' &&
      result !== null &&
      'error' in result &&
      typeof result.error === 'string'
        ? result.error
        : JSON.stringify(result);
    throw new Error(`Hachidori dictionary upload failed (${response.status}): ${detail}`);
  }
}
