import assert from 'node:assert/strict';
import test from 'node:test';

import { describeDownloadError } from './utils.js';

test('describeDownloadError prefers the error message', () => {
  assert.equal(describeDownloadError(new Error('socket hang up')), 'socket hang up');
});

test('describeDownloadError falls back to the error code when the message is empty', () => {
  const err = new Error('') as NodeJS.ErrnoException;
  err.code = 'ECONNRESET';
  assert.equal(describeDownloadError(err), 'ECONNRESET');
});

test('describeDownloadError unwraps empty-message AggregateErrors', () => {
  const v4 = new Error('connect ECONNREFUSED 1.2.3.4:443') as NodeJS.ErrnoException;
  v4.code = 'ECONNREFUSED';
  const v6 = new Error('') as NodeJS.ErrnoException;
  v6.code = 'ENETUNREACH';
  const aggregate = new AggregateError([v4, v6], '');
  assert.equal(describeDownloadError(aggregate), 'connect ECONNREFUSED 1.2.3.4:443; ENETUNREACH');
});

test('describeDownloadError never returns an empty string', () => {
  assert.equal(describeDownloadError(new Error('')), 'Error');
  assert.equal(describeDownloadError('boom'), 'boom');
});
