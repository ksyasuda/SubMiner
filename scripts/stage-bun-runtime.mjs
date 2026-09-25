import { createHash, randomUUID } from 'node:crypto';
import { createReadStream, createWriteStream } from 'node:fs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { pipeline } from 'node:stream/promises';
import { Readable } from 'node:stream';
import { promisify } from 'node:util';
import { execFile } from 'node:child_process';

const execFileAsync = promisify(execFile);
const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, '..');

export const DEFAULT_PACKAGE_JSON_PATH = path.join(repoRoot, 'package.json');
export const DEFAULT_MANIFEST_PATH = path.join(repoRoot, 'build', 'bun-runtime-manifest.json');
export const DEFAULT_CACHE_DIR = path.join(repoRoot, '.tmp', 'bun-runtime');
export const DEFAULT_LICENSES_SOURCE_DIR = path.join(repoRoot, 'resources', 'bun', 'licenses');
export const STAGED_METADATA_FILE = 'metadata.json';
export const REQUIRED_LICENSE_FILES = [
  'Bun-LICENSE.md',
  'LGPL-2.0.txt',
  'LGPL-2.1.txt',
  'SOURCE.md',
  'THIRD-PARTY-NOTICES.md',
];

const RELEASE_BASE_URL = 'https://github.com/oven-sh/bun/releases/download';
const SUPPORTED_PLATFORMS = new Set(['darwin', 'linux', 'win32']);
const ARCH_BY_BUILDER_VALUE = new Map([
  [1, 'x64'],
  [3, 'arm64'],
  ['x64', 'x64'],
  ['arm64', 'arm64'],
]);

function isRecord(value) {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function readRequiredString(record, key, source) {
  const value = record[key];
  if (typeof value !== 'string' || value.length === 0) {
    throw new Error(`${source} must contain a non-empty ${key} string.`);
  }
  return value;
}

function readCommit(record, key, source) {
  const value = readRequiredString(record, key, source);
  if (!/^[a-f0-9]{40}$/.test(value)) {
    throw new Error(`${source} must contain a 40-character ${key} commit.`);
  }
  return value;
}

export function normalizeTarget(platform, arch) {
  if (!SUPPORTED_PLATFORMS.has(platform)) {
    throw new Error(`Unsupported Bun runtime target platform: ${platform}`);
  }

  const normalizedArch = ARCH_BY_BUILDER_VALUE.get(arch);
  if (!normalizedArch) {
    throw new Error(`Unsupported Bun runtime target architecture for ${platform}: ${String(arch)}`);
  }
  if (platform === 'win32' && normalizedArch !== 'x64') {
    throw new Error(`Unsupported Bun runtime target: ${platform}-${normalizedArch}`);
  }

  return {
    platform,
    arch: normalizedArch,
    key: `${platform}-${normalizedArch}`,
    executableName: platform === 'win32' ? 'bun.exe' : 'bun',
  };
}

export function parsePackageManagerVersion(packageJson) {
  if (!isRecord(packageJson)) {
    throw new Error('package.json must contain a JSON object.');
  }
  const packageManager = readRequiredString(packageJson, 'packageManager', 'package.json');
  const match = /^bun@(\d+\.\d+\.\d+)$/.exec(packageManager);
  if (!match) {
    throw new Error(
      `package.json packageManager must pin Bun exactly, received ${packageManager}.`,
    );
  }
  return match[1];
}

function parseArtifact(value, key) {
  if (!isRecord(value)) {
    throw new Error(`Bun runtime manifest artifact ${key} must be an object.`);
  }
  const file = readRequiredString(value, 'file', `Bun runtime manifest artifact ${key}`);
  const sha256 = readRequiredString(value, 'sha256', `Bun runtime manifest artifact ${key}`);
  if (!/^[a-f0-9]{64}$/.test(sha256)) {
    throw new Error(`Bun runtime manifest artifact ${key} has an invalid SHA-256 digest.`);
  }
  if (path.basename(file) !== file || !file.endsWith('.zip')) {
    throw new Error(`Bun runtime manifest artifact ${key} has an unsafe file name.`);
  }
  return { file, sha256 };
}

export function parseRuntimeManifest(manifest, version) {
  if (!isRecord(manifest) || manifest.schemaVersion !== 1) {
    throw new Error('Bun runtime manifest must use schemaVersion 1.');
  }
  const manifestVersion = readRequiredString(manifest, 'version', 'Bun runtime manifest');
  if (manifestVersion !== version) {
    throw new Error(
      `Bun runtime manifest version ${manifestVersion} does not match packageManager bun@${version}.`,
    );
  }
  if (!isRecord(manifest.artifacts)) {
    throw new Error('Bun runtime manifest must contain an artifacts object.');
  }
  return {
    version,
    bunRevision: readCommit(manifest, 'bunRevision', 'Bun runtime manifest'),
    releaseTagCommit: readCommit(manifest, 'releaseTagCommit', 'Bun runtime manifest'),
    artifacts: manifest.artifacts,
    licenseInventoryStatus: readRequiredString(
      manifest,
      'licenseInventoryStatus',
      'Bun runtime manifest',
    ),
    sourceManifest: readRequiredString(manifest, 'sourceManifest', 'Bun runtime manifest'),
    correspondingSourceAsset: readRequiredString(
      manifest,
      'correspondingSourceAsset',
      'Bun runtime manifest',
    ),
  };
}

export async function loadRuntimeConfig({
  packageJsonPath = DEFAULT_PACKAGE_JSON_PATH,
  manifestPath = DEFAULT_MANIFEST_PATH,
} = {}) {
  const [packageJsonText, manifestText] = await Promise.all([
    fs.readFile(packageJsonPath, 'utf8'),
    fs.readFile(manifestPath, 'utf8'),
  ]);
  const version = parsePackageManagerVersion(JSON.parse(packageJsonText));
  return parseRuntimeManifest(JSON.parse(manifestText), version);
}

export function resolveArtifact(config, platform, arch) {
  const target = normalizeTarget(platform, arch);
  const artifactValue = config.artifacts[target.key];
  if (artifactValue === undefined) {
    throw new Error(`Bun runtime manifest has no artifact for ${target.key}.`);
  }
  const artifact = parseArtifact(artifactValue, target.key);
  return {
    ...target,
    ...artifact,
    version: config.version,
    url: `${RELEASE_BASE_URL}/bun-v${config.version}/${artifact.file}`,
  };
}

export async function sha256File(filePath) {
  const hash = createHash('sha256');
  for await (const chunk of createReadStream(filePath)) {
    hash.update(chunk);
  }
  return hash.digest('hex');
}

export async function ensureCachedArchive(
  artifact,
  { cacheDir = DEFAULT_CACHE_DIR, fetchImpl = globalThis.fetch } = {},
) {
  if (typeof fetchImpl !== 'function') {
    throw new Error('No fetch implementation is available to download Bun.');
  }
  const versionCacheDir = path.join(cacheDir, artifact.version);
  const archivePath = path.join(versionCacheDir, artifact.file);
  await fs.mkdir(versionCacheDir, { recursive: true });

  try {
    if ((await sha256File(archivePath)) === artifact.sha256) return archivePath;
    await fs.unlink(archivePath);
  } catch (error) {
    if (!isRecord(error) || error.code !== 'ENOENT') throw error;
  }

  const temporaryPath = `${archivePath}.download-${randomUUID()}`;
  try {
    const response = await fetchImpl(artifact.url);
    if (!response.ok || !response.body) {
      throw new Error(`Unable to download ${artifact.url}: HTTP ${response.status}`);
    }
    await pipeline(
      Readable.fromWeb(response.body),
      createWriteStream(temporaryPath, { flags: 'wx' }),
    );
    const actualSha256 = await sha256File(temporaryPath);
    if (actualSha256 !== artifact.sha256) {
      throw new Error(
        `Bun archive checksum mismatch for ${artifact.file}: expected ${artifact.sha256}, received ${actualSha256}.`,
      );
    }
    await fs.rename(temporaryPath, archivePath);
    return archivePath;
  } catch (error) {
    await fs.rm(temporaryPath, { force: true });
    throw error;
  }
}

function isSafeZipEntry(entry) {
  if (entry.startsWith('/') || /^[A-Za-z]:/.test(entry)) return false;
  return !entry.replaceAll('\\', '/').split('/').includes('..');
}

function waitForProcess(child, description) {
  let stderr = '';
  child.stderr.setEncoding('utf8');
  child.stderr.on('data', (chunk) => {
    stderr += chunk;
  });
  return new Promise((resolve, reject) => {
    child.once('error', reject);
    child.once('close', (code) => {
      if (code === 0) resolve();
      else reject(new Error(`${description}: ${stderr.trim() || `process exited ${code}`}`));
    });
  });
}

export function buildWindowsExtractionCommand(archivePath, member, outputPath) {
  const script = [
    'Add-Type -AssemblyName System.IO.Compression.FileSystem',
    '$archive = [IO.Compression.ZipFile]::OpenRead($env:SUBMINER_BUN_ARCHIVE_PATH)',
    'try {',
    '  $entry = $archive.Entries | Where-Object { $_.FullName -ceq $env:SUBMINER_BUN_ARCHIVE_MEMBER }',
    '  if ($null -eq $entry) { throw "Archive member not found: $env:SUBMINER_BUN_ARCHIVE_MEMBER" }',
    '  $inputStream = $entry.Open()',
    '  $outputStream = [IO.File]::Create($env:SUBMINER_BUN_OUTPUT_PATH)',
    '  try { $inputStream.CopyTo($outputStream) } finally { $outputStream.Dispose(); $inputStream.Dispose() }',
    '} finally { $archive.Dispose() }',
  ].join('; ');
  return {
    command: 'powershell.exe',
    args: [
      '-NoLogo',
      '-NoProfile',
      '-NonInteractive',
      '-EncodedCommand',
      Buffer.from(script, 'utf16le').toString('base64'),
    ],
    environment: {
      SUBMINER_BUN_ARCHIVE_PATH: archivePath,
      SUBMINER_BUN_ARCHIVE_MEMBER: member,
      SUBMINER_BUN_OUTPUT_PATH: outputPath,
    },
  };
}

async function extractZipMemberOnWindows(archivePath, member, outputPath) {
  const command = buildWindowsExtractionCommand(archivePath, member, outputPath);
  const powershell = spawn(command.command, command.args, {
    env: { ...process.env, ...command.environment },
    stdio: ['ignore', 'ignore', 'pipe'],
  });
  await waitForProcess(powershell, `Unable to extract ${member}`);
}

export async function extractZipMember(archivePath, member, outputPath) {
  if (!isSafeZipEntry(member)) {
    throw new Error(`Refusing to extract unsafe Bun archive member ${member}.`);
  }
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const temporaryPath = `${outputPath}.extract-${randomUUID()}`;

  if (process.platform === 'win32') {
    try {
      await extractZipMemberOnWindows(archivePath, member, temporaryPath);
      await fs.chmod(temporaryPath, 0o755);
      await fs.rename(temporaryPath, outputPath);
    } catch (error) {
      await fs.rm(temporaryPath, { force: true });
      throw error;
    }
    return;
  }

  const { stdout } = await execFileAsync('unzip', ['-Z1', archivePath], {
    encoding: 'utf8',
    maxBuffer: 1024 * 1024,
  });
  const entries = stdout.split(/\r?\n/).filter(Boolean);
  if (entries.some((entry) => !isSafeZipEntry(entry))) {
    throw new Error(`Bun archive ${archivePath} contains an unsafe path.`);
  }
  if (!entries.includes(member)) {
    throw new Error(`Bun archive ${archivePath} does not contain ${member}.`);
  }

  const output = createWriteStream(temporaryPath, { flags: 'wx', mode: 0o755 });
  const unzip = spawn('unzip', ['-p', archivePath, member], { stdio: ['ignore', 'pipe', 'pipe'] });

  try {
    await Promise.all([
      pipeline(unzip.stdout, output),
      waitForProcess(unzip, `Unable to extract ${member}`),
    ]);
    await fs.chmod(temporaryPath, 0o755);
    await fs.rename(temporaryPath, outputPath);
  } catch (error) {
    await fs.rm(temporaryPath, { force: true });
    throw error;
  }
}

export function resolveResourcesDirectory(appOutDir, platform, productFilename = 'SubMiner') {
  if (platform !== 'darwin') return path.join(appOutDir, 'resources');
  const appBundlePath = appOutDir.endsWith('.app')
    ? appOutDir
    : path.join(appOutDir, `${productFilename}.app`);
  return path.join(appBundlePath, 'Contents', 'Resources');
}

export async function stageBunLicenses(
  runtimeDirectory,
  licensesSourceDir = DEFAULT_LICENSES_SOURCE_DIR,
) {
  const licensesDirectory = path.join(runtimeDirectory, 'licenses');
  await Promise.all(
    REQUIRED_LICENSE_FILES.map((fileName) => fs.access(path.join(licensesSourceDir, fileName))),
  );
  await fs.cp(licensesSourceDir, licensesDirectory, { recursive: true, force: true });
  return licensesDirectory;
}

export async function stageBunRuntime(
  { appOutDir, platform, arch, productFilename = 'SubMiner' },
  {
    configLoader = loadRuntimeConfig,
    archiveLoader = ensureCachedArchive,
    extractor = extractZipMember,
    licenseStager = stageBunLicenses,
  } = {},
) {
  const config = await configLoader();
  const artifact = resolveArtifact(config, platform, arch);
  const archivePath = await archiveLoader(artifact);
  const archiveDirectory = path.basename(artifact.file, '.zip');
  const archiveExecutableName = platform === 'win32' ? 'bun.exe' : 'bun';
  const member = `${archiveDirectory}/${archiveExecutableName}`;
  const runtimeDirectory = path.join(
    resolveResourcesDirectory(appOutDir, platform, productFilename),
    'bun',
  );
  const executablePath = path.join(runtimeDirectory, artifact.executableName);
  await extractor(archivePath, member, executablePath);
  await licenseStager(runtimeDirectory);

  const metadata = {
    name: 'Bun',
    version: config.version,
    bunRevision: config.bunRevision,
    releaseTagCommit: config.releaseTagCommit,
    target: artifact.key,
    artifact: artifact.file,
    artifactSha256: artifact.sha256,
    sourceUrl: artifact.url,
    licenseInventoryStatus: config.licenseInventoryStatus,
    correspondingSourceAsset: config.correspondingSourceAsset,
    sourceInstructions: 'licenses/SOURCE.md',
    thirdPartyNotices: 'licenses/THIRD-PARTY-NOTICES.md',
  };
  await fs.writeFile(
    path.join(runtimeDirectory, STAGED_METADATA_FILE),
    `${JSON.stringify(metadata, null, 2)}\n`,
    { mode: 0o644 },
  );
  return {
    executablePath,
    metadataPath: path.join(runtimeDirectory, STAGED_METADATA_FILE),
    artifact,
  };
}
