import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';
import type { HachidoriHostStatus } from '../../../shared/hachidori-sharing';

// hachidori-docker names itself this in the sharing hello; browser and app hosts
// report their own name and have no management API to upload to.
export const HACHIDORI_DOCKER_HOST_NAME = 'Hachidori Docker host';
export const HACHIDORI_DOCKER_MANAGEMENT_PORT = 8780;

const LOOPBACK_HOSTNAMES = new Set(['127.0.0.1', 'localhost', '[::1]']);

function sameMachine(left: string, right: string): boolean {
  return left === right || (LOOPBACK_HOSTNAMES.has(left) && LOOPBACK_HOSTNAMES.has(right));
}

// Picks the management origin that holds the linked host's dictionaries. The link
// address names the machine; `override` (hachidori.externalHostManagementUrl, already
// normalized to an origin) covers a non-default port or a reverse proxy.
export function resolveHachidoriManagementUrl(
  host: Pick<Extract<HachidoriHostStatus, { kind: 'connected' }>, 'address' | 'name'>,
  override: string,
  zipPath: string,
  warn?: (message: string) => void,
): string {
  const linkHostname = new URL(host.address).hostname;
  if (override) {
    const overrideHostname = new URL(override).hostname;
    if (!sameMachine(overrideHostname, linkHostname)) {
      warn?.(
        `hachidori.externalHostManagementUrl (${override}) points at ${overrideHostname}, but Hachidori is linked to ${linkHostname}; uploads may miss the linked host.`,
      );
    }
    return override;
  }
  if (host.name !== HACHIDORI_DOCKER_HOST_NAME) {
    throw new Error(
      `The linked ${host.name} at ${linkHostname} cannot receive dictionary uploads. Import ${zipPath} from its Hachidori settings, or link a Hachidori Docker host.`,
    );
  }
  return new URL(`http://${linkHostname}:${HACHIDORI_DOCKER_MANAGEMENT_PORT}`).origin;
}

// Upload from the main process: extension blob URLs cannot cross the sharing link,
// and the Docker management API deliberately rejects browser cross-origin writes.
export async function uploadHachidoriDictionary(zipPath: string, origin: string): Promise<void> {
  const url = new URL('/import', origin);
  url.searchParams.set('name', path.basename(zipPath));
  url.searchParams.set('replace', 'true');
  const bytes = await readFile(zipPath);
  const signal = AbortSignal.timeout(300_000);
  for (;;) {
    let response: Response;
    try {
      response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/zip' },
        body: bytes,
        signal,
        redirect: 'error',
      });
    } catch (error) {
      // undici reports every network failure as "fetch failed"; the cause says why.
      const cause = error instanceof Error && error.cause instanceof Error ? error.cause : error;
      const reason = cause instanceof Error ? cause.message : String(cause);
      throw new Error(`Could not reach the Hachidori management API at ${origin}: ${reason}`);
    }
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
