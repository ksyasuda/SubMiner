import assert from 'node:assert/strict';
import test from 'node:test';

import { mergeLcovReports } from './run-coverage-lane';

test('mergeLcovReports combines duplicate source-file counters across shard outputs', () => {
  const merged = mergeLcovReports([
    [
      'SF:src/example.ts',
      'FN:10,alpha',
      'FNDA:1,alpha',
      'DA:10,1',
      'DA:11,0',
      'BRDA:10,0,0,1',
      'BRDA:10,0,1,-',
      'end_of_record',
      '',
    ].join('\n'),
    [
      'SF:src/example.ts',
      'FN:10,alpha',
      'FN:20,beta',
      'FNDA:2,alpha',
      'FNDA:1,beta',
      'DA:10,2',
      'DA:11,1',
      'DA:20,1',
      'BRDA:10,0,0,0',
      'BRDA:10,0,1,1',
      'end_of_record',
      '',
    ].join('\n'),
  ]);

  assert.match(merged, /SF:src\/example\.ts/);
  assert.match(merged, /FN:10,alpha/);
  assert.match(merged, /FN:20,beta/);
  assert.match(merged, /FNDA:3,alpha/);
  assert.match(merged, /FNDA:1,beta/);
  assert.match(merged, /FNF:2/);
  assert.match(merged, /FNH:2/);
  assert.match(merged, /DA:10,3/);
  assert.match(merged, /DA:11,1/);
  assert.match(merged, /DA:20,1/);
  assert.match(merged, /LF:3/);
  assert.match(merged, /LH:3/);
  assert.match(merged, /BRDA:10,0,0,1/);
  assert.match(merged, /BRDA:10,0,1,1/);
  assert.match(merged, /BRF:2/);
  assert.match(merged, /BRH:2/);
});

test('mergeLcovReports keeps distinct source files as separate records', () => {
  const merged = mergeLcovReports([
    ['SF:src/a.ts', 'DA:1,1', 'end_of_record', ''].join('\n'),
    ['SF:src/b.ts', 'DA:2,1', 'end_of_record', ''].join('\n'),
  ]);

  assert.match(merged, /SF:src\/a\.ts[\s\S]*end_of_record/);
  assert.match(merged, /SF:src\/b\.ts[\s\S]*end_of_record/);
});
