import test from 'node:test';
import assert from 'node:assert/strict';
import { buildSubminerScriptOpts } from './script-opts';

test('buildSubminerScriptOpts preserves app and socket paths verbatim', () => {
  const scriptOpts = buildSubminerScriptOpts(
    '/Applications/SubMiner  Beta.app/Contents/MacOS/SubMiner',
    '/tmp/subminer  socket.sock',
    ['subminer-backend=x11'],
  );

  assert.equal(
    scriptOpts,
    'subminer-binary_path=/Applications/SubMiner  Beta.app/Contents/MacOS/SubMiner,subminer-socket_path=/tmp/subminer  socket.sock,subminer-backend=x11',
  );
});

test('buildSubminerScriptOpts rejects delimiter-bearing default paths', () => {
  assert.throws(
    () => buildSubminerScriptOpts('/tmp/SubMiner,canary', '/tmp/subminer.sock'),
    /subminer-binary_path contains unsupported script option delimiter/,
  );
  assert.throws(
    () => buildSubminerScriptOpts('/tmp/SubMiner', '/tmp/subminer\nsocket.sock'),
    /subminer-socket_path contains unsupported script option delimiter/,
  );
});
