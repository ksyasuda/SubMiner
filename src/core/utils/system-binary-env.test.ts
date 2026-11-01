import assert from 'node:assert/strict';
import test from 'node:test';

import { buildSystemBinaryEnv } from './system-binary-env';

test('system binary environment drops the AppImage library-path override', () => {
  const env = buildSystemBinaryEnv({
    LD_LIBRARY_PATH: '/tmp/.mount_SubMinXXXXXX/usr/lib',
    DBUS_SESSION_BUS_ADDRESS: 'unix:path=/run/user/1000/bus',
    HOME: '/home/user',
  });

  assert.equal(env.LD_LIBRARY_PATH, undefined);
  assert.equal(env.DBUS_SESSION_BUS_ADDRESS, 'unix:path=/run/user/1000/bus');
  assert.equal(env.HOME, '/home/user');
});
