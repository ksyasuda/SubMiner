import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import {
  buildWindowsExtractionCommand,
  ensureCachedArchive,
  extractZipMember,
  loadRuntimeConfig,
  normalizeTarget,
  parsePackageManagerVersion,
  parseRuntimeManifest,
  resolveArtifact,
  stageBunLicenses,
  stageBunRuntime,
} from './stage-bun-runtime.mjs';

function sha256(content: Uint8Array): string {
  return createHash('sha256').update(content).digest('hex');
}

function runtimeConfig() {
  return {
    version: '1.3.5',
    bunRevision: '1e86cebd74a5723e818b5c0555276b646bcf0e4c',
    releaseTagCommit: 'fa5a5bbe556a4bda5bde77b4013aa6c3bb4ec9ab',
    artifacts: {
      'darwin-arm64': { file: 'bun-darwin-aarch64.zip', sha256: '1'.repeat(64) },
      'darwin-x64': { file: 'bun-darwin-x64-baseline.zip', sha256: '2'.repeat(64) },
      'linux-arm64': { file: 'bun-linux-aarch64.zip', sha256: '3'.repeat(64) },
      'linux-x64': { file: 'bun-linux-x64-baseline.zip', sha256: '4'.repeat(64) },
      'win32-x64': { file: 'bun-windows-x64-baseline.zip', sha256: '5'.repeat(64) },
    },
    licenseInventoryStatus: 'complete-for-bun-1.3.5-declared-linked-libraries',
    sourceManifest: 'build/bun-source-manifest.json',
    correspondingSourceAsset: 'bun-v1.3.5-source.tar.gz',
  };
}

test('normalizeTarget maps each supported electron-builder target without using the host', () => {
  assert.deepEqual(normalizeTarget('linux', 1), {
    platform: 'linux',
    arch: 'x64',
    key: 'linux-x64',
    executableName: 'bun',
  });
  assert.equal(normalizeTarget('linux', 3).key, 'linux-arm64');
  assert.equal(normalizeTarget('darwin', 'x64').key, 'darwin-x64');
  assert.equal(normalizeTarget('darwin', 'arm64').key, 'darwin-arm64');
  assert.deepEqual(normalizeTarget('win32', 1), {
    platform: 'win32',
    arch: 'x64',
    key: 'win32-x64',
    executableName: 'bun.exe',
  });
  assert.throws(() => normalizeTarget('freebsd', 'x64'), /Unsupported Bun runtime target platform/);
  assert.throws(() => normalizeTarget('win32', 'arm64'), /Unsupported Bun runtime target/);
  assert.throws(() => normalizeTarget('linux', 0), /Unsupported Bun runtime target architecture/);
});

test('resolveArtifact chooses baseline x64 builds and standard arm64 builds', () => {
  const config = runtimeConfig();
  assert.equal(resolveArtifact(config, 'linux', 'x64').file, 'bun-linux-x64-baseline.zip');
  assert.equal(resolveArtifact(config, 'darwin', 'x64').file, 'bun-darwin-x64-baseline.zip');
  assert.equal(resolveArtifact(config, 'win32', 'x64').file, 'bun-windows-x64-baseline.zip');
  assert.equal(resolveArtifact(config, 'linux', 'arm64').file, 'bun-linux-aarch64.zip');
  assert.equal(resolveArtifact(config, 'darwin', 'arm64').file, 'bun-darwin-aarch64.zip');
});

test('tracked runtime manifest covers every supported target at the packageManager version', async () => {
  const config = await loadRuntimeConfig();
  assert.equal(config.version, '1.3.5');
  for (const [platform, arch] of [
    ['linux', 'x64'],
    ['linux', 'arm64'],
    ['darwin', 'x64'],
    ['darwin', 'arm64'],
    ['win32', 'x64'],
  ]) {
    const artifact = resolveArtifact(config, platform, arch);
    assert.match(artifact.sha256, /^[a-f0-9]{64}$/);
    assert.equal(
      artifact.url,
      `https://github.com/oven-sh/bun/releases/download/bun-v1.3.5/${artifact.file}`,
    );
  }
});

test('manifest version must match the exact packageManager Bun pin', () => {
  assert.equal(parsePackageManagerVersion({ packageManager: 'bun@1.3.5' }), '1.3.5');
  assert.throws(
    () => parsePackageManagerVersion({ packageManager: 'bun@^1.3.5' }),
    /must pin Bun exactly/,
  );
  assert.throws(
    () =>
      parseRuntimeManifest(
        {
          schemaVersion: 1,
          version: '1.3.6',
          bunRevision: '1e86cebd74a5723e818b5c0555276b646bcf0e4c',
          releaseTagCommit: 'fa5a5bbe556a4bda5bde77b4013aa6c3bb4ec9ab',
          artifacts: {},
          licenseInventoryStatus: 'partial',
          sourceManifest: 'build/bun-source-manifest.json',
          correspondingSourceAsset: 'bun-v1.3.5-source.tar.gz',
        },
        '1.3.5',
      ),
    /does not match packageManager bun@1\.3\.5/,
  );
});

test('ensureCachedArchive verifies downloads and reuses a verified cache entry offline', async () => {
  const workspace = await fs.mkdtemp(path.join(os.tmpdir(), 'subminer-bun-cache-'));
  const bytes = new TextEncoder().encode('verified Bun archive fixture');
  const artifact = {
    version: '1.3.5',
    file: 'bun-linux-x64-baseline.zip',
    sha256: sha256(bytes),
    url: 'https://example.invalid/bun.zip',
  };
  let downloads = 0;
  const fetchImpl = async () => {
    downloads += 1;
    return new Response(bytes);
  };

  try {
    const firstPath = await ensureCachedArchive(artifact, { cacheDir: workspace, fetchImpl });
    const secondPath = await ensureCachedArchive(artifact, {
      cacheDir: workspace,
      fetchImpl: async () => {
        throw new Error('verified cache should not fetch');
      },
    });
    assert.equal(firstPath, secondPath);
    assert.equal(downloads, 1);
    assert.deepEqual(await fs.readFile(firstPath), Buffer.from(bytes));
  } finally {
    await fs.rm(workspace, { recursive: true, force: true });
  }
});

test('ensureCachedArchive rejects a checksum mismatch without caching the download', async () => {
  const workspace = await fs.mkdtemp(path.join(os.tmpdir(), 'subminer-bun-cache-mismatch-'));
  const artifact = {
    version: '1.3.5',
    file: 'bun-linux-x64-baseline.zip',
    sha256: '0'.repeat(64),
    url: 'https://example.invalid/bun.zip',
  };

  try {
    await assert.rejects(
      ensureCachedArchive(artifact, {
        cacheDir: workspace,
        fetchImpl: async () => new Response('tampered'),
      }),
      /checksum mismatch/,
    );
    const cacheEntries = await fs.readdir(path.join(workspace, '1.3.5'));
    assert.deepEqual(cacheEntries, []);
  } finally {
    await fs.rm(workspace, { recursive: true, force: true });
  }
});

test('extractZipMember extracts only the requested path and makes the runtime executable', async () => {
  const workspace = await fs.mkdtemp(path.join(os.tmpdir(), 'subminer-bun-extract-'));
  const archivePath = path.join(workspace, 'fixture.zip');
  const outputPath = path.join(workspace, 'output', 'bun');
  const archiveBase64 =
    'UEsDBAoAAAAAAMxcKl3rs337EgAAABIAAAAaABwAYnVuLWxpbnV4LXg2NC1iYXNlbGluZS9idW5VVAkAAyD5omog+aJqdXgLAAEE6AMAAAToAwAAYnVuIGZpeHR1cmUgYmluYXJ5UEsBAh4DCgAAAAAAzFwqXeuzffsSAAAAEgAAABoAGAAAAAAAAQAAAKSBAAAAAGJ1bi1saW51eC14NjQtYmFzZWxpbmUvYnVuVVQFAAMg+aJqdXgLAAEE6AMAAAToAwAAUEsFBgAAAAABAAEAYAAAAGYAAAAAAA==';

  try {
    await fs.writeFile(archivePath, Buffer.from(archiveBase64, 'base64'));
    await extractZipMember(archivePath, 'bun-linux-x64-baseline/bun', outputPath);
    assert.equal(await fs.readFile(outputPath, 'utf8'), 'bun fixture binary');
    assert.equal((await fs.stat(outputPath)).mode & 0o777, 0o755);
    await assert.rejects(
      extractZipMember(archivePath, '../bun', path.join(workspace, 'unsafe')),
      /unsafe Bun archive member/,
    );
  } finally {
    await fs.rm(workspace, { recursive: true, force: true });
  }
});

test('Windows extraction keeps paths and archive members out of PowerShell command text', () => {
  const archivePath = String.raw`C:\Release Builds\bun $(archive) '1.3.5'.zip`;
  const member = 'bun-windows-x64-baseline/bun.exe';
  const outputPath = String.raw`C:\Staged App & Tools\resources\bun\bun.exe`;
  const command = buildWindowsExtractionCommand(archivePath, member, outputPath);
  const encodedCommand = command.args.at(-1);

  assert.equal(command.command, 'powershell.exe');
  assert.equal(command.args.at(-2), '-EncodedCommand');
  assert.ok(encodedCommand);
  const script = Buffer.from(encodedCommand, 'base64').toString('utf16le');
  assert.match(script, /\$env:SUBMINER_BUN_ARCHIVE_PATH/);
  assert.match(script, /\$env:SUBMINER_BUN_ARCHIVE_MEMBER/);
  assert.match(script, /\$env:SUBMINER_BUN_OUTPUT_PATH/);
  assert.doesNotMatch(script, /Release Builds|Staged App|bun-windows-x64-baseline/);
  assert.deepEqual(command.environment, {
    SUBMINER_BUN_ARCHIVE_PATH: archivePath,
    SUBMINER_BUN_ARCHIVE_MEMBER: member,
    SUBMINER_BUN_OUTPUT_PATH: outputPath,
  });
});

test('stageBunLicenses copies the tracked inventory into resources/bun/licenses', async () => {
  const workspace = await fs.mkdtemp(path.join(os.tmpdir(), 'subminer-bun-licenses-'));
  const sourceDirectory = path.join(workspace, 'tracked-licenses');
  const runtimeDirectory = path.join(workspace, 'resources', 'bun');

  try {
    await fs.mkdir(sourceDirectory, { recursive: true });
    for (const fileName of [
      'Bun-LICENSE.md',
      'LGPL-2.0.txt',
      'LGPL-2.1.txt',
      'SOURCE.md',
      'THIRD-PARTY-NOTICES.md',
    ]) {
      await fs.writeFile(path.join(sourceDirectory, fileName), `${fileName}\n`);
    }
    const licensesDirectory = await stageBunLicenses(runtimeDirectory, sourceDirectory);
    assert.equal(licensesDirectory, path.join(runtimeDirectory, 'licenses'));
    assert.equal(
      await fs.readFile(path.join(licensesDirectory, 'THIRD-PARTY-NOTICES.md'), 'utf8'),
      'THIRD-PARTY-NOTICES.md\n',
    );
  } finally {
    await fs.rm(workspace, { recursive: true, force: true });
  }
});

test('stageBunRuntime places target executables and metadata in app resources with safe modes', async () => {
  const workspace = await fs.mkdtemp(path.join(os.tmpdir(), 'subminer-bun-stage-'));
  const cases = [
    {
      platform: 'linux',
      arch: 'x64',
      appOutDir: path.join(workspace, 'linux'),
      relativeExecutable: path.join('resources', 'bun', 'bun'),
      member: 'bun-linux-x64-baseline/bun',
    },
    {
      platform: 'darwin',
      arch: 'arm64',
      appOutDir: path.join(workspace, 'darwin'),
      relativeExecutable: path.join('SubMiner.app', 'Contents', 'Resources', 'bun', 'bun'),
      member: 'bun-darwin-aarch64/bun',
    },
    {
      platform: 'win32',
      arch: 'x64',
      appOutDir: path.join(workspace, 'windows'),
      relativeExecutable: path.join('resources', 'bun', 'bun.exe'),
      member: 'bun-windows-x64-baseline/bun.exe',
    },
  ] as const;

  try {
    for (const targetCase of cases) {
      let extractedMember = '';
      const result = await stageBunRuntime(targetCase, {
        configLoader: async () => runtimeConfig(),
        archiveLoader: async () => path.join(workspace, 'fixture.zip'),
        extractor: async (_archivePath: string, member: string, outputPath: string) => {
          extractedMember = member;
          await fs.mkdir(path.dirname(outputPath), { recursive: true });
          await fs.writeFile(outputPath, 'bun fixture', { mode: 0o755 });
        },
        licenseStager: async (runtimeDirectory: string) => {
          const licensesDirectory = path.join(runtimeDirectory, 'licenses');
          await fs.mkdir(licensesDirectory, { recursive: true });
          await fs.writeFile(path.join(licensesDirectory, 'Bun-LICENSE.md'), 'MIT fixture');
          return licensesDirectory;
        },
      });
      const expectedExecutable = path.join(targetCase.appOutDir, targetCase.relativeExecutable);
      assert.equal(result.executablePath, expectedExecutable);
      assert.equal(extractedMember, targetCase.member);
      assert.equal((await fs.stat(expectedExecutable)).mode & 0o777, 0o755);
      assert.equal((await fs.stat(result.metadataPath)).mode & 0o777, 0o644);
      const metadata = JSON.parse(await fs.readFile(result.metadataPath, 'utf8'));
      assert.equal(metadata.version, '1.3.5');
      assert.equal(metadata.bunRevision, '1e86cebd74a5723e818b5c0555276b646bcf0e4c');
      assert.equal(metadata.artifactSha256, result.artifact.sha256);
      assert.equal(metadata.target, result.artifact.key);
      assert.equal(
        await fs.readFile(
          path.join(path.dirname(expectedExecutable), 'licenses', 'Bun-LICENSE.md'),
          'utf8',
        ),
        'MIT fixture',
      );
    }
  } finally {
    await fs.rm(workspace, { recursive: true, force: true });
  }
});
