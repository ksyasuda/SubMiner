import test from 'node:test';
import assert from 'node:assert/strict';
import { createUpdateDialogPresenter, type ShowMessageBox } from './update-dialogs';

test('update dialog presenter focuses app before showing macOS dialogs', async () => {
  const calls: string[] = [];
  const showMessageBox: ShowMessageBox = async (options) => {
    calls.push(`dialog:${options.message}`);
    return { response: 0 };
  };
  const presenter = createUpdateDialogPresenter({
    platform: 'darwin',
    focusApp: () => calls.push('focus'),
    showMessageBox,
  });

  await presenter.showNoUpdateDialog('0.14.0');

  assert.deepEqual(calls, ['focus', 'dialog:SubMiner is up to date (v0.14.0)']);
});

test('update dialog presenter does not focus app before showing non-macOS dialogs', async () => {
  const calls: string[] = [];
  const showMessageBox: ShowMessageBox = async (options) => {
    calls.push(`dialog:${options.message}`);
    return { response: 0 };
  };
  const presenter = createUpdateDialogPresenter({
    platform: 'linux',
    focusApp: () => calls.push('focus'),
    showMessageBox,
  });

  await presenter.showNoUpdateDialog('0.14.0');

  assert.deepEqual(calls, ['dialog:SubMiner is up to date (v0.14.0)']);
});
