import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createUpdateDialogPresenter,
  showManualUpdateRequiredDialog,
  type ShowMessageBox,
} from './update-dialogs';

test('update dialog presenter focuses app and yields the run loop before showing macOS dialogs', async () => {
  const calls: string[] = [];
  const showMessageBox: ShowMessageBox = async (options) => {
    calls.push(`dialog:${options.message}`);
    return { response: 0 };
  };
  const presenter = createUpdateDialogPresenter({
    platform: 'darwin',
    focusApp: () => {
      calls.push('focus');
    },
    yieldToRunLoop: async () => {
      calls.push('yield');
    },
    showMessageBox,
  });

  await presenter.showNoUpdateDialog('0.14.0');

  assert.deepEqual(calls, ['focus', 'yield', 'dialog:SubMiner is up to date (v0.14.0)']);
});

test('update dialog presenter awaits async focusApp before yielding and showing the dialog', async () => {
  const calls: string[] = [];
  const showMessageBox: ShowMessageBox = async (options) => {
    calls.push(`dialog:${options.message}`);
    return { response: 0 };
  };
  const presenter = createUpdateDialogPresenter({
    platform: 'darwin',
    focusApp: async () => {
      await new Promise<void>((resolve) => setImmediate(resolve));
      calls.push('focus');
    },
    yieldToRunLoop: async () => {
      calls.push('yield');
    },
    showMessageBox,
  });

  await presenter.showNoUpdateDialog('0.14.0');

  assert.deepEqual(calls, ['focus', 'yield', 'dialog:SubMiner is up to date (v0.14.0)']);
});

test('update dialog presenter does not focus app or yield before showing non-macOS dialogs', async () => {
  const calls: string[] = [];
  const showMessageBox: ShowMessageBox = async (options) => {
    calls.push(`dialog:${options.message}`);
    return { response: 0 };
  };
  const presenter = createUpdateDialogPresenter({
    platform: 'linux',
    focusApp: () => {
      calls.push('focus');
    },
    yieldToRunLoop: async () => {
      calls.push('yield');
    },
    showMessageBox,
  });

  await presenter.showNoUpdateDialog('0.14.0');

  assert.deepEqual(calls, ['dialog:SubMiner is up to date (v0.14.0)']);
});

test('manual update required dialog explains that automatic install is unavailable', async () => {
  let shown:
    | {
        type?: string;
        title?: string;
        message: string;
        detail?: string;
        buttons?: string[];
      }
    | undefined;
  const showMessageBox: ShowMessageBox = async (options) => {
    shown = options;
    return { response: 0 };
  };

  await showManualUpdateRequiredDialog(showMessageBox, '0.15.0-beta.1');

  assert.equal(shown?.type, 'warning');
  assert.equal(shown?.message, 'Manual install required');
  assert.match(shown?.detail ?? '', /SubMiner v0\.15\.0-beta\.1 is available/);
  assert.match(shown?.detail ?? '', /cannot install app updates automatically/);
});
