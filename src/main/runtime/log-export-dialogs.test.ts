import assert from 'node:assert/strict';
import test from 'node:test';
import { showLogExportErrorDialog, showLogExportSuccessDialog } from './log-export-dialogs';

test('showLogExportSuccessDialog handles dialog rejection', async () => {
  const warnings: string[] = [];

  await showLogExportSuccessDialog({
    zipPath: '/tmp/subminer-logs.zip',
    showMessageBox: async () => {
      throw new Error('dialog failed');
    },
    showItemInFolder: () => {
      throw new Error('unexpected shell call');
    },
    logWarn: (message) => warnings.push(message),
  });

  assert.deepEqual(warnings, ['Failed to show log export success dialog.']);
});

test('showLogExportErrorDialog handles dialog rejection', async () => {
  const warnings: string[] = [];

  await showLogExportErrorDialog({
    message: 'export failed',
    showMessageBox: async () => {
      throw new Error('dialog failed');
    },
    logWarn: (message) => warnings.push(message),
  });

  assert.deepEqual(warnings, ['Failed to show log export error dialog.']);
});
