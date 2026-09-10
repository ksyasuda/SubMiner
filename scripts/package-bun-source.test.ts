import { describe, expect, test } from 'bun:test';
import fs from 'node:fs/promises';
import path from 'node:path';
import {
  parseRegisteredRepositories,
  parseSourceManifest,
  validateBunPins,
  validateRuntimeAlignment,
} from './package-bun-source.mjs';

const projectRoot = path.resolve(import.meta.dir, '..');

describe('Bun corresponding-source manifest', () => {
  test('pins the source to the revision reported by the distributed binary', async () => {
    const manifest = parseSourceManifest(
      JSON.parse(
        await fs.readFile(path.join(projectRoot, 'build/bun-source-manifest.json'), 'utf8'),
      ),
    );

    expect(manifest.version).toBe('1.3.5');
    expect(manifest.bunRevision).toBe('1e86cebd74a5723e818b5c0555276b646bcf0e4c');
    expect(manifest.releaseTagCommit).toBe('fa5a5bbe556a4bda5bde77b4013aa6c3bb4ec9ab');
    expect(manifest.sources.find((source) => source.name === 'WebKit')?.revision).toBe(
      '6d0f3aac0b817cc01a846b3754b21271adedac12',
    );
    expect(manifest.sources.find((source) => source.name === 'tinycc')?.revision).toBe(
      '29985a3b59898861442fa3b43f663fc1af2591d7',
    );
  });

  test('parses commit and tag registrations from Bun CMake', () => {
    const registrations = parseRegisteredRepositories(`
      register_repository(
        NAME tinycc
        REPOSITORY oven-sh/tinycc
        COMMIT
          # A comment between the field and its value is valid CMake.
          29985a3b59898861442fa3b43f663fc1af2591d7
      )
      register_repository(
        NAME brotli
        REPOSITORY google/brotli
        TAG v1.1.0
      )
    `);

    expect(registrations.get('tinycc')).toEqual({
      repository: 'oven-sh/tinycc',
      kind: 'commit',
      reference: '29985a3b59898861442fa3b43f663fc1af2591d7',
    });
    expect(registrations.get('brotli')).toEqual({
      repository: 'google/brotli',
      kind: 'tag',
      reference: 'v1.1.0',
    });
  });

  test('rejects stale runtime or package-manager pins', async () => {
    const manifest = parseSourceManifest(
      JSON.parse(
        await fs.readFile(path.join(projectRoot, 'build/bun-source-manifest.json'), 'utf8'),
      ),
    );
    const runtimeManifest = {
      version: manifest.version,
      bunRevision: manifest.bunRevision,
      correspondingSourceAsset: manifest.archiveName,
    };

    expect(() =>
      validateRuntimeAlignment(manifest, { packageManager: 'bun@1.3.5' }, runtimeManifest),
    ).not.toThrow();
    expect(() =>
      validateRuntimeAlignment(manifest, { packageManager: 'bun@1.3.6' }, runtimeManifest),
    ).toThrow('package.json must pin bun@1.3.5');
    expect(() =>
      validateRuntimeAlignment(
        manifest,
        { packageManager: 'bun@1.3.5' },
        {
          ...runtimeManifest,
          correspondingSourceAsset: 'stale.tar.gz',
        },
      ),
    ).toThrow('correspondingSourceAsset does not match');
  });

  test('rejects drift in CMake dependency and WebKit pins', async () => {
    const manifest = parseSourceManifest(
      JSON.parse(
        await fs.readFile(path.join(projectRoot, 'build/bun-source-manifest.json'), 'utf8'),
      ),
    );
    const registrations = manifest.sources
      .filter((source) => source.destination.startsWith('bun/vendor/') && source.name !== 'WebKit')
      .map(
        (source) => `register_repository(
          NAME ${source.name}
          REPOSITORY ${source.repository}
          ${source.upstreamReference ? 'TAG' : 'COMMIT'} ${source.upstreamReference ?? source.revision}
        )`,
      )
      .join('\n');

    expect(() =>
      validateBunPins(
        manifest,
        [registrations],
        'set(WEBKIT_VERSION 6d0f3aac0b817cc01a846b3754b21271adedac12)',
      ),
    ).not.toThrow();
    expect(() =>
      validateBunPins(
        manifest,
        [registrations.replace('29985a3b59898861442fa3b43f663fc1af2591d7', '0'.repeat(40))],
        'set(WEBKIT_VERSION 6d0f3aac0b817cc01a846b3754b21271adedac12)',
      ),
    ).toThrow('revision mismatch for tinycc');
    expect(() =>
      validateBunPins(manifest, [registrations], `set(WEBKIT_VERSION ${'0'.repeat(40)})`),
    ).toThrow('WebKit revision does not match');
  });
});
