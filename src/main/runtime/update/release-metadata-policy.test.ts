import assert from 'node:assert/strict';
import test from 'node:test';
import { shouldFetchReleaseMetadataForPlatform } from './release-metadata-policy';

test('macOS automatic release metadata fetch is skipped when native updater is unsupported', () => {
  assert.equal(
    shouldFetchReleaseMetadataForPlatform('darwin', {
      available: false,
      version: '0.14.0',
      canUpdate: false,
    }),
    false,
  );
  assert.equal(
    shouldFetchReleaseMetadataForPlatform('darwin', {
      available: false,
      version: '0.14.0',
    }),
    true,
  );
  assert.equal(
    shouldFetchReleaseMetadataForPlatform('darwin', {
      available: true,
      version: '0.15.0',
      canUpdate: true,
    }),
    true,
  );
});

test('macOS manual checks fetch release metadata when native updater is unsupported', () => {
  const unsupportedUpdate = {
    available: false,
    version: '0.15.0-beta.4',
    canUpdate: false,
  };

  assert.equal(
    shouldFetchReleaseMetadataForPlatform('darwin', unsupportedUpdate, {
      source: 'manual',
    }),
    true,
  );
  assert.equal(
    shouldFetchReleaseMetadataForPlatform('darwin', unsupportedUpdate, {
      source: 'launcher',
    }),
    true,
  );
  assert.equal(
    shouldFetchReleaseMetadataForPlatform('darwin', unsupportedUpdate, {
      source: 'automatic',
    }),
    false,
  );
});

test('non-macOS release metadata fetch is not gated by native updater support', () => {
  assert.equal(
    shouldFetchReleaseMetadataForPlatform('linux', {
      available: false,
      version: '0.14.0',
      canUpdate: false,
    }),
    true,
  );
  assert.equal(
    shouldFetchReleaseMetadataForPlatform('win32', {
      available: false,
      version: '0.14.0',
      canUpdate: false,
    }),
    true,
  );
});
