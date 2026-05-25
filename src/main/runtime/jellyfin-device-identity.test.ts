import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createHostDerivedJellyfinDeviceId,
  resolveJellyfinRemoteDeviceName,
} from './jellyfin-device-identity';

test('createHostDerivedJellyfinDeviceId uses the hostname as the stable id', () => {
  assert.equal(createHostDerivedJellyfinDeviceId('Kyle-PC.local'), 'Kyle-PC.local');
  assert.equal(createHostDerivedJellyfinDeviceId(''), 'device');
});

test('resolveJellyfinRemoteDeviceName uses hostname by default', () => {
  assert.equal(
    resolveJellyfinRemoteDeviceName({
      hostName: 'kyle-pc',
    }),
    'kyle-pc',
  );
});

test('resolveJellyfinRemoteDeviceName falls back when hostname is empty', () => {
  assert.equal(resolveJellyfinRemoteDeviceName({ hostName: '' }), 'device');
});
