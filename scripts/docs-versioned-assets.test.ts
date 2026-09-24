import { describe, expect, test } from 'bun:test';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { dedupeVersionedPublicAssets, rewriteSharedAssetReferences } from './docs-versioned-assets';

function tempDir() {
  return mkdtempSync(join(tmpdir(), 'subminer-docs-versioned-assets-'));
}

describe('docs versioned asset dedupe', () => {
  test('rewrites version-scoped public asset references to shared root assets', () => {
    const html =
      '<link href="/v/0.14.0/assets/style.hash.css"><source src="/v/0.14.0/assets/minecard.webm">';
    const expected =
      '<link href="/v/0.14.0/assets/style.hash.css"><source src="/assets/minecard.webm">';

    expect(rewriteSharedAssetReferences(html, '/v/0.14.0/', new Set(['minecard.webm']))).toBe(
      expected,
    );
    expect(rewriteSharedAssetReferences(html, '/v/0.14.0', new Set(['minecard.webm']))).toBe(
      expected,
    );
  });

  test('does not rewrite longer asset paths with a shared asset prefix', () => {
    expect(
      rewriteSharedAssetReferences(
        [
          '<script src="/v/0.14.0/assets/foo.js"></script>',
          '<script src="/v/0.14.0/assets/foo.js?v=1"></script>',
          '<script src="/v/0.14.0/assets/foo.js.map"></script>',
        ].join(''),
        '/v/0.14.0/',
        new Set(['foo.js']),
      ),
    ).toBe(
      [
        '<script src="/assets/foo.js"></script>',
        '<script src="/assets/foo.js?v=1"></script>',
        '<script src="/v/0.14.0/assets/foo.js.map"></script>',
      ].join(''),
    );
  });

  test('removes duplicated version public assets while preserving generated VitePress assets', async () => {
    const dir = tempDir();
    try {
      mkdirSync(join(dir, 'assets/chunks'), { recursive: true });
      writeFileSync(join(dir, 'assets/style.hash.css'), 'body{}');
      writeFileSync(join(dir, 'assets/chunks/theme.hash.js'), 'export {};');
      writeFileSync(join(dir, 'assets/minecard.webm'), 'large video');
      writeFileSync(
        join(dir, 'index.html'),
        '<link href="/v/0.14.0/assets/style.hash.css"><source src="/v/0.14.0/assets/minecard.webm">',
      );

      const result = dedupeVersionedPublicAssets({
        outDir: dir,
        base: '/v/0.14.0',
        sharedAssetPaths: new Set(['minecard.webm']),
      });

      expect(result.rewrittenFiles).toEqual([join(dir, 'index.html')]);
      expect(existsSync(join(dir, 'assets/style.hash.css'))).toBe(true);
      expect(existsSync(join(dir, 'assets/chunks/theme.hash.js'))).toBe(true);
      expect(existsSync(join(dir, 'assets/minecard.webm'))).toBe(false);
      expect(readFileSync(join(dir, 'index.html'), 'utf8')).toBe(
        '<link href="/v/0.14.0/assets/style.hash.css"><source src="/assets/minecard.webm">',
      );
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  test('keeps root public assets because they are the shared copy', async () => {
    const dir = tempDir();
    try {
      mkdirSync(join(dir, 'assets'), { recursive: true });
      writeFileSync(join(dir, 'assets/minecard.webm'), 'large video');

      const result = dedupeVersionedPublicAssets({
        outDir: dir,
        base: '/',
        sharedAssetPaths: new Set(['minecard.webm']),
      });

      expect(result.removedAssetsDir).toBe(false);
      expect(existsSync(join(dir, 'assets'))).toBe(true);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
