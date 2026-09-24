import { spawnSync } from 'node:child_process';
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { versionOutputPath } from './docs-versioning';

// Frozen `/v/<version>/` doc builds live in an R2 bucket and are served by the Pages
// Function in `docs-site/functions/v/[[path]].ts`, so they never count toward the Pages
// deployment. Transfers go through the AWS CLI's S3 API (preinstalled on GitHub runners).

// Written last on upload; an archive without it is treated as missing and rebuilt.
export const ARCHIVE_MARKER = '_archive.json';

export type DocsArchiveStore = {
  has(version: string): boolean;
  upload(version: string, dir: string, builtFrom: string): void;
};

export type DocsArchiveStoreEnv = {
  accountId?: string;
  accessKeyId?: string;
  secretAccessKey?: string;
  bucket?: string;
};

export function archiveStoreEnvFromProcess(env: NodeJS.ProcessEnv): DocsArchiveStoreEnv {
  return {
    accountId: env.CLOUDFLARE_ACCOUNT_ID,
    accessKeyId: env.DOCS_ARCHIVE_R2_ACCESS_KEY_ID,
    secretAccessKey: env.DOCS_ARCHIVE_R2_SECRET_ACCESS_KEY,
    bucket: env.DOCS_ARCHIVE_R2_BUCKET,
  };
}

export function archiveKeyPrefix(version: string): string {
  return `${versionOutputPath(version)}/`;
}

// Returns null when credentials are absent so local builds can skip archive sync.
export function createArchiveStore(config: DocsArchiveStoreEnv): DocsArchiveStore | null {
  const { accountId, accessKeyId, secretAccessKey, bucket } = config;
  if (!accountId || !accessKeyId || !secretAccessKey || !bucket) {
    return null;
  }

  const endpoint = `https://${accountId}.r2.cloudflarestorage.com`;
  const env: NodeJS.ProcessEnv = {
    ...process.env,
    AWS_ACCESS_KEY_ID: accessKeyId,
    AWS_SECRET_ACCESS_KEY: secretAccessKey,
    AWS_DEFAULT_REGION: 'auto',
    // AWS CLI >= 2.23 sends CRC checksums by default, which R2 does not fully support.
    AWS_REQUEST_CHECKSUM_CALCULATION: 'when_required',
    AWS_RESPONSE_CHECKSUM_VALIDATION: 'when_required',
  };

  function aws(args: string[]) {
    return spawnSync('aws', [...args, '--endpoint-url', endpoint], {
      env,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    });
  }

  return {
    has(version) {
      const result = aws([
        's3api',
        'head-object',
        '--bucket',
        bucket,
        '--key',
        `${archiveKeyPrefix(version)}${ARCHIVE_MARKER}`,
      ]);
      if (result.error) throw result.error;
      if (result.status === 0) return true;
      if (/\b404\b|Not Found/i.test(result.stderr)) return false;
      throw new Error(`Unable to check docs archive ${version}: ${result.stderr.trim()}`);
    },

    upload(version, dir, builtFrom) {
      const target = `s3://${bucket}/${archiveKeyPrefix(version)}`;
      const sync = aws(['s3', 'sync', dir, target, '--only-show-errors']);
      if (sync.error) throw sync.error;
      if (sync.status !== 0) {
        throw new Error(`Unable to upload docs archive ${version}: ${sync.stderr.trim()}`);
      }

      const markerPath = join(dir, ARCHIVE_MARKER);
      writeFileSync(
        markerPath,
        `${JSON.stringify({ version, builtFrom, builtAt: new Date().toISOString() }, null, 2)}\n`,
      );
      const marker = aws(['s3', 'cp', markerPath, `${target}${ARCHIVE_MARKER}`]);
      if (marker.status !== 0) {
        throw new Error(`Unable to mark docs archive ${version}: ${marker.stderr.trim()}`);
      }
    },
  };
}
