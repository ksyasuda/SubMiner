import { createHash, randomUUID } from 'node:crypto';
import { createReadStream, createWriteStream } from 'node:fs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { pipeline } from 'node:stream/promises';
import { Readable } from 'node:stream';

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, '..');

export const DEFAULT_MANIFEST_PATH = path.join(repoRoot, 'build', 'bun-source-manifest.json');
export const DEFAULT_RUNTIME_MANIFEST_PATH = path.join(
  repoRoot,
  'build',
  'bun-runtime-manifest.json',
);
export const DEFAULT_PACKAGE_JSON_PATH = path.join(repoRoot, 'package.json');
export const DEFAULT_OUTPUT_DIR = path.join(repoRoot, 'release');
export const DEFAULT_CACHE_DIR = path.join(repoRoot, '.tmp', 'bun-corresponding-source');

function isRecord(value) {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function requiredString(record, key, source) {
  const value = record[key];
  if (typeof value !== 'string' || value.length === 0) {
    throw new Error(`${source} must contain a non-empty ${key} string.`);
  }
  return value;
}

function assertSafeRelativePath(value, source) {
  if (path.isAbsolute(value) || value.split(/[\\/]/).includes('..')) {
    throw new Error(`${source} contains an unsafe path: ${value}`);
  }
}

export function parseSourceManifest(value) {
  if (!isRecord(value) || value.schemaVersion !== 1) {
    throw new Error('Bun source manifest must use schemaVersion 1.');
  }
  const version = requiredString(value, 'version', 'Bun source manifest');
  const bunRevision = requiredString(value, 'bunRevision', 'Bun source manifest');
  const archiveName = requiredString(value, 'archiveName', 'Bun source manifest');
  if (!/^\d+\.\d+\.\d+$/.test(version) || !/^[a-f0-9]{40}$/.test(bunRevision)) {
    throw new Error('Bun source manifest has an invalid version or bunRevision.');
  }
  if (path.basename(archiveName) !== archiveName || !archiveName.endsWith('.tar.gz')) {
    throw new Error('Bun source manifest archiveName must be a safe .tar.gz filename.');
  }
  if (!Array.isArray(value.sources) || value.sources.length === 0) {
    throw new Error('Bun source manifest must contain sources.');
  }

  const names = new Set();
  const destinations = new Set();
  const sources = value.sources.map((entry, index) => {
    if (!isRecord(entry)) throw new Error(`Bun source ${index} must be an object.`);
    const name = requiredString(entry, 'name', `Bun source ${index}`);
    const repository = requiredString(entry, 'repository', `Bun source ${name}`);
    const revision = requiredString(entry, 'revision', `Bun source ${name}`);
    const destination = requiredString(entry, 'destination', `Bun source ${name}`);
    if (!/^[\w.-]+\/[\w.-]+$/.test(repository) || !/^[a-f0-9]{40}$/.test(revision)) {
      throw new Error(`Bun source ${name} has an invalid repository or revision.`);
    }
    assertSafeRelativePath(destination, `Bun source ${name}`);
    if (names.has(name) || destinations.has(destination)) {
      throw new Error(`Bun source manifest repeats ${name} or ${destination}.`);
    }
    names.add(name);
    destinations.add(destination);

    const transport = entry.transport ?? 'archive';
    if (transport !== 'archive' && transport !== 'git-sparse') {
      throw new Error(`Bun source ${name} has unsupported transport ${transport}.`);
    }
    const sha256 =
      transport === 'archive' ? requiredString(entry, 'sha256', `Bun source ${name}`) : null;
    if (sha256 !== null && !/^[a-f0-9]{64}$/.test(sha256)) {
      throw new Error(`Bun source ${name} has an invalid SHA-256 digest.`);
    }
    if (!Array.isArray(entry.licensePaths) || entry.licensePaths.length === 0) {
      throw new Error(`Bun source ${name} must declare licensePaths.`);
    }
    const licensePaths = entry.licensePaths.map((licensePath) => {
      if (typeof licensePath !== 'string' || licensePath.length === 0) {
        throw new Error(`Bun source ${name} has an invalid license path.`);
      }
      assertSafeRelativePath(licensePath, `Bun source ${name}`);
      return licensePath;
    });
    const exclude = Array.isArray(entry.exclude) ? entry.exclude : [];
    for (const excludedPath of exclude) assertSafeRelativePath(excludedPath, `Bun source ${name}`);
    return {
      ...entry,
      name,
      repository,
      revision,
      destination,
      transport,
      sha256,
      licensePaths,
      exclude,
    };
  });

  const bun = sources.find((source) => source.name === 'bun');
  if (!bun || bun.revision !== bunRevision || bun.destination !== 'bun') {
    throw new Error('Bun source manifest must map bunRevision to the bun source at bun/.');
  }
  return { ...value, version, bunRevision, archiveName, sources };
}

export function parseRegisteredRepositories(cmakeText) {
  const registrations = new Map();
  const uncommented = cmakeText.replace(/#[^\n]*/g, '');
  for (const match of uncommented.matchAll(/register_repository\(([\s\S]*?)\)/g)) {
    const body = match[1];
    const name = /\bNAME\s+([^\s#)]+)/.exec(body)?.[1];
    const repository = /\bREPOSITORY\s+([^\s#)]+)/.exec(body)?.[1];
    const reference = /\b(COMMIT|TAG)\s+(?:#[^\n]*\n\s*)?([^\s#)]+)/.exec(body);
    if (name && repository && reference) {
      registrations.set(name, {
        repository,
        kind: reference[1].toLowerCase(),
        reference: reference[2],
      });
    }
  }
  return registrations;
}

export function validateRuntimeAlignment(manifest, packageJson, runtimeManifest) {
  if (!isRecord(packageJson) || packageJson.packageManager !== `bun@${manifest.version}`) {
    throw new Error(
      `package.json must pin bun@${manifest.version} to match the Bun source manifest.`,
    );
  }
  if (!isRecord(runtimeManifest)) throw new Error('Bun runtime manifest must be an object.');
  for (const key of ['version', 'bunRevision']) {
    if (runtimeManifest[key] !== manifest[key]) {
      throw new Error(`Bun runtime manifest ${key} does not match the Bun source manifest.`);
    }
  }
  if (runtimeManifest.correspondingSourceAsset !== manifest.archiveName) {
    throw new Error(
      'Bun runtime manifest correspondingSourceAsset does not match the Bun source manifest.',
    );
  }
}

export function validateBunPins(manifest, cmakeTexts, setupWebKitText) {
  const actual = new Map();
  for (const cmakeText of cmakeTexts) {
    for (const [name, registration] of parseRegisteredRepositories(cmakeText)) {
      if (actual.has(name)) throw new Error(`Bun registers ${name} more than once.`);
      actual.set(name, registration);
    }
  }
  const expected = new Map(
    manifest.sources
      .filter((source) => source.destination.startsWith('bun/vendor/') && source.name !== 'WebKit')
      .map((source) => [source.name, source]),
  );
  for (const [name, registration] of actual) {
    const source = expected.get(name);
    if (!source) throw new Error(`Source manifest omits Bun dependency ${name}.`);
    if (source.repository !== registration.repository) {
      throw new Error(`Source manifest repository mismatch for ${name}.`);
    }
    const expectedReference = source.upstreamReference ?? source.revision;
    if (expectedReference !== registration.reference) {
      throw new Error(`Source manifest revision mismatch for ${name}.`);
    }
    expected.delete(name);
  }
  if (expected.size > 0) {
    throw new Error(
      `Source manifest has unregistered Bun dependencies: ${[...expected.keys()].join(', ')}.`,
    );
  }

  const webKitPin = /set\(WEBKIT_VERSION\s+([a-f0-9]{40})\)/.exec(setupWebKitText)?.[1];
  const webKit = manifest.sources.find((source) => source.name === 'WebKit');
  if (!webKitPin || !webKit || webKit.revision !== webKitPin) {
    throw new Error('Source manifest WebKit revision does not match SetupWebKit.cmake.');
  }
}

async function sha256File(filePath) {
  const hash = createHash('sha256');
  for await (const chunk of createReadStream(filePath)) hash.update(chunk);
  return hash.digest('hex');
}

async function run(command, args, options = {}) {
  await new Promise((resolve, reject) => {
    const child = spawn(command, args, { stdio: 'inherit', ...options });
    child.once('error', reject);
    child.once('close', (code) => {
      if (code === 0) resolve();
      else reject(new Error(`${command} exited with status ${code}.`));
    });
  });
}

async function downloadArchive(source, cacheDir, fetchImpl) {
  const archivePath = path.join(cacheDir, `${source.name}-${source.revision}.tar.gz`);
  try {
    if ((await sha256File(archivePath)) === source.sha256) return archivePath;
    await fs.rm(archivePath, { force: true });
  } catch (error) {
    if (!isRecord(error) || error.code !== 'ENOENT') throw error;
  }

  const url = `https://codeload.github.com/${source.repository}/tar.gz/${source.revision}`;
  const temporaryPath = `${archivePath}.download-${randomUUID()}`;
  const response = await fetchImpl(url);
  if (!response.ok || !response.body)
    throw new Error(`Unable to download ${url}: HTTP ${response.status}`);
  await pipeline(
    Readable.fromWeb(response.body),
    createWriteStream(temporaryPath, { flags: 'wx' }),
  );
  const actualSha256 = await sha256File(temporaryPath);
  if (actualSha256 !== source.sha256) {
    await fs.rm(temporaryPath, { force: true });
    throw new Error(
      `Source checksum mismatch for ${source.name}: expected ${source.sha256}, received ${actualSha256}.`,
    );
  }
  await fs.rename(temporaryPath, archivePath);
  return archivePath;
}

async function materializeArchive(source, root, cacheDir, fetchImpl) {
  const archivePath = await downloadArchive(source, cacheDir, fetchImpl);
  const destination = path.join(root, source.destination);
  await fs.mkdir(destination, { recursive: true });
  await run('tar', ['-xzf', archivePath, '-C', destination, '--strip-components=1']);
}

async function materializeGitSparse(source, root, cacheDir) {
  const checkout = path.join(cacheDir, `${source.name}-${source.revision}-git`);
  await fs.rm(checkout, { recursive: true, force: true });
  await run('git', [
    'clone',
    '--filter=blob:none',
    '--no-checkout',
    '--depth=1',
    `https://github.com/${source.repository}.git`,
    checkout,
  ]);
  await run('git', ['-C', checkout, 'fetch', '--depth=1', 'origin', source.revision]);
  await run('git', ['-C', checkout, 'sparse-checkout', 'init', '--no-cone']);
  const sparseRules = ['/*', ...source.exclude.map((entry) => `!/${entry}/`), ''];
  await fs.writeFile(
    path.join(checkout, '.git', 'info', 'sparse-checkout'),
    sparseRules.join('\n'),
  );
  await run('git', ['-C', checkout, 'checkout', '--detach', source.revision]);
  const actualRevision = (await fs.readFile(path.join(checkout, '.git', 'HEAD'), 'utf8')).trim();
  if (actualRevision !== source.revision)
    throw new Error(`Git checkout mismatch for ${source.name}.`);
  await fs.rm(path.join(checkout, '.git'), { recursive: true, force: true });
  await fs.mkdir(path.dirname(path.join(root, source.destination)), { recursive: true });
  await fs.rename(checkout, path.join(root, source.destination));
}

async function applyBunDependencyPatches(root, manifest) {
  const bunRoot = path.join(root, 'bun');
  for (const source of manifest.sources) {
    if (!source.destination.startsWith('bun/vendor/') || source.name === 'WebKit') continue;
    const destination = path.join(root, source.destination);
    const patchDirectory = path.join(bunRoot, 'patches', source.name);
    let entries = [];
    try {
      entries = await fs.readdir(patchDirectory, { withFileTypes: true });
    } catch (error) {
      if (!isRecord(error) || error.code !== 'ENOENT') throw error;
    }
    for (const entry of entries.sort((left, right) => left.name.localeCompare(right.name))) {
      const patchPath = path.join(patchDirectory, entry.name);
      if (entry.isFile() && entry.name.endsWith('.patch')) {
        await run(
          'git',
          ['apply', '--ignore-whitespace', '--ignore-space-change', '--no-index', patchPath],
          { cwd: destination },
        );
      } else if (entry.isFile()) {
        await fs.copyFile(patchPath, path.join(destination, entry.name));
      }
    }
    const cmakeReference = source.upstreamReference
      ? `refs/tags/${source.upstreamReference}`
      : source.revision;
    await fs.writeFile(path.join(destination, '.ref'), `${cmakeReference}\n`);
  }
}

async function validateAndCollectLicenses(root, manifest) {
  const licensesRoot = path.join(root, 'THIRD-PARTY-LICENSES');
  await fs.mkdir(licensesRoot, { recursive: true });
  for (const source of manifest.sources) {
    const target = path.join(licensesRoot, source.name);
    await fs.mkdir(target, { recursive: true });
    for (const licensePath of source.licensePaths) {
      const sourcePath = path.join(root, source.destination, licensePath);
      const stat = await fs.stat(sourcePath).catch(() => null);
      if (!stat?.isFile() || stat.size === 0) {
        throw new Error(`Missing required license material for ${source.name}: ${licensePath}`);
      }
      const safeName = licensePath.replaceAll('/', '__');
      await fs.copyFile(sourcePath, path.join(target, safeName));
    }
  }
}

function rebuildReadme(manifest) {
  const webKit = manifest.sources.find((source) => source.name === 'WebKit');
  const tinycc = manifest.sources.find((source) => source.name === 'tinycc');
  return `# Bun ${manifest.version} corresponding source and rebuild materials

This archive matches the official Bun ${manifest.version} binaries whose \`bun --revision\` output names commit \`${manifest.bunRevision}\`. The GitHub release tag points to \`${manifest.releaseTagCommit}\`, one later commit, so this package intentionally uses the binary revision.

The archive includes Bun's complete source tree, Bun's build scripts and dependency patches, every external repository registered by Bun's CMake build at its exact pin, and Oven's WebKit fork at \`${webKit.revision}\`. Large WebKit test-only trees are excluded. The JavaScriptCore, WTF, WebCore, build-tool, configuration, and resource trees used to build the library are included. TinyCC is \`${tinycc.revision}\`.

## Rebuild with modified JavaScriptCore

Install the prerequisites recorded in \`bun/.buildkite/Dockerfile\`, \`bun/scripts/bootstrap.sh\`, and the WebKit platform build scripts. Bun ${manifest.version} used LLVM 19.1.7, CMake 3.30.5 in its Linux build image, and Bun 1.1.38 as the bootstrap runtime. Its Rust input was nightly and was not pinned to a dated toolchain in the release source.

From this archive root on Linux or macOS:

\`\`\`sh
cd bun
bun install --frozen-lockfile
bun run jsc:build
bun run build:release:local -- -DVERSION=${manifest.version} -DREVISION=${manifest.bunRevision}
\`\`\`

The first command uses the bootstrap Bun. \`jsc:build\` builds the included \`vendor/WebKit\` checkout into \`vendor/WebKit/WebKitBuild/Release\`. \`build:release:local\` links Bun against that local JavaScriptCore build. The explicit version and revision replace metadata that Bun normally reads from its Git checkout. The included \`vendor/*/.ref\` files prevent Bun's CMake rules from replacing the packaged dependency sources, and this package has already applied the files under \`bun/patches/<dependency>/\` in the same order as \`bun/cmake/scripts/GitClone.cmake\`.

The archive vendors the source repositories that Bun's CMake build links into the executable. It preserves Bun's \`bun.lock\` files and lol-html's \`Cargo.lock\`, but it does not vendor npm packages, crates.io packages used to build lol-html, compilers, SDKs, or other build tools. The rebuild therefore needs network access for those pinned package-manager inputs. License notices embedded in those downloaded packages are outside the collected \`THIRD-PARTY-LICENSES\` directory's scope.

Windows uses the prerequisites in \`bun/docs/project/building-windows.mdx\` and WebKit's \`windows-release.ps1\`. The local-JavaScriptCore path above has not been verified on Windows.

These instructions describe the source and build entry points. Toolchain and generated-output differences mean a rebuild is not expected to be byte-for-byte identical to Oven's release binary. No claim about legal compliance or reproducible builds is made here.
`;
}

async function createDeterministicArchive(stagingParent, rootName, outputPath) {
  const temporaryTar = `${outputPath}.tar-${randomUUID()}`;
  const temporaryGzip = `${outputPath}.gzip-${randomUUID()}`;
  try {
    await run('tar', [
      '--sort=name',
      '--mtime=@0',
      '--owner=0',
      '--group=0',
      '--numeric-owner',
      '-cf',
      temporaryTar,
      '-C',
      stagingParent,
      rootName,
    ]);
    const gzip = spawn('gzip', ['-n', '-9', '-c', temporaryTar], {
      stdio: ['ignore', 'pipe', 'inherit'],
    });
    const completion = new Promise((resolve, reject) => {
      gzip.once('error', reject);
      gzip.once('close', resolve);
    });
    const [, code] = await Promise.all([
      pipeline(gzip.stdout, createWriteStream(temporaryGzip, { flags: 'wx' })),
      completion,
    ]);
    if (code !== 0) throw new Error(`gzip exited with status ${code}.`);
    await fs.rm(outputPath, { force: true });
    await fs.rename(temporaryGzip, outputPath);
  } finally {
    await Promise.all([
      fs.rm(temporaryTar, { force: true }),
      fs.rm(temporaryGzip, { force: true }),
    ]);
  }
}

export async function packageBunSource({
  manifestPath = DEFAULT_MANIFEST_PATH,
  runtimeManifestPath = DEFAULT_RUNTIME_MANIFEST_PATH,
  packageJsonPath = DEFAULT_PACKAGE_JSON_PATH,
  outputDir = DEFAULT_OUTPUT_DIR,
  cacheDir = DEFAULT_CACHE_DIR,
  fetchImpl = globalThis.fetch,
} = {}) {
  const [manifestText, runtimeManifestText, packageJsonText] = await Promise.all([
    fs.readFile(manifestPath, 'utf8'),
    fs.readFile(runtimeManifestPath, 'utf8'),
    fs.readFile(packageJsonPath, 'utf8'),
  ]);
  const manifest = parseSourceManifest(JSON.parse(manifestText));
  validateRuntimeAlignment(manifest, JSON.parse(packageJsonText), JSON.parse(runtimeManifestText));
  if (typeof fetchImpl !== 'function') throw new Error('No fetch implementation is available.');
  await fs.mkdir(cacheDir, { recursive: true });
  const stagingParent = await fs.mkdtemp(path.join(cacheDir, 'assemble-'));
  const rootName = path.basename(manifest.archiveName, '.tar.gz');
  const root = path.join(stagingParent, rootName);
  await fs.mkdir(root);

  try {
    const bun = manifest.sources.find((source) => source.name === 'bun');
    await materializeArchive(bun, root, cacheDir, fetchImpl);
    const cmakeFiles = (await fs.readdir(path.join(root, 'bun', 'cmake', 'targets')))
      .filter((name) => name.endsWith('.cmake'))
      .map((name) => fs.readFile(path.join(root, 'bun', 'cmake', 'targets', name), 'utf8'));
    validateBunPins(
      manifest,
      await Promise.all(cmakeFiles),
      await fs.readFile(path.join(root, 'bun', 'cmake', 'tools', 'SetupWebKit.cmake'), 'utf8'),
    );

    for (const source of manifest.sources.filter((entry) => entry.name !== 'bun')) {
      if (source.transport === 'git-sparse') await materializeGitSparse(source, root, cacheDir);
      else await materializeArchive(source, root, cacheDir, fetchImpl);
    }
    await applyBunDependencyPatches(root, manifest);
    await validateAndCollectLicenses(root, manifest);
    await fs.writeFile(
      path.join(root, 'SOURCE-INVENTORY.json'),
      `${JSON.stringify(manifest, null, 2)}\n`,
    );
    await fs.writeFile(path.join(root, 'README-REBUILD.md'), rebuildReadme(manifest));

    await fs.mkdir(outputDir, { recursive: true });
    const outputPath = path.join(outputDir, manifest.archiveName);
    await createDeterministicArchive(stagingParent, rootName, outputPath);
    const digest = await sha256File(outputPath);
    await fs.writeFile(`${outputPath}.sha256`, `${digest}  ${manifest.archiveName}\n`);
    return { outputPath, sha256: digest };
  } finally {
    await fs.rm(stagingParent, { recursive: true, force: true });
  }
}

if (import.meta.main) {
  const result = await packageBunSource();
  console.log(`${result.sha256}  ${path.basename(result.outputPath)}`);
}
