import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { runAppCommandCaptureOutput } from '../mpv.js';
import { nowMs } from '../time.js';
import { sleep } from '../util.js';
import type { LauncherCommandContext } from './context.js';

type UpdateCommandResponse = {
  ok: boolean;
  status?: string;
  version?: string;
  error?: string;
};

type UpdateCommandDeps = {
  createTempDir: (prefix: string) => string;
  joinPath: (...parts: string[]) => string;
  runAppCommandCaptureOutput: (
    appPath: string,
    appArgs: string[],
  ) => { status: number; stdout: string; stderr: string; error?: Error };
  waitForUpdateResponse: (responsePath: string) => Promise<UpdateCommandResponse>;
  removeDir: (targetPath: string) => void;
};

const UPDATE_RESPONSE_TIMEOUT_MS = 10 * 60 * 1000;

const defaultDeps: UpdateCommandDeps = {
  createTempDir: (prefix) => fs.mkdtempSync(path.join(os.tmpdir(), prefix)),
  joinPath: (...parts) => path.join(...parts),
  runAppCommandCaptureOutput: (appPath, appArgs) => runAppCommandCaptureOutput(appPath, appArgs),
  waitForUpdateResponse: async (responsePath) => {
    const deadline = nowMs() + UPDATE_RESPONSE_TIMEOUT_MS;
    while (nowMs() < deadline) {
      try {
        if (fs.existsSync(responsePath)) {
          return JSON.parse(fs.readFileSync(responsePath, 'utf8')) as UpdateCommandResponse;
        }
      } catch {
        // retry until timeout
      }
      await sleep(100);
    }
    return { ok: false, error: 'Timed out waiting for SubMiner update response.' };
  },
  removeDir: (targetPath) => {
    fs.rmSync(targetPath, { recursive: true, force: true });
  },
};

export async function runUpdateCommand(
  context: LauncherCommandContext,
  deps: Partial<UpdateCommandDeps> = {},
): Promise<boolean> {
  const resolvedDeps: UpdateCommandDeps = { ...defaultDeps, ...deps };
  const { args, appPath, scriptPath } = context;
  if (!args.update || !appPath) {
    return false;
  }

  const tempDir = resolvedDeps.createTempDir('subminer-update-');
  const responsePath = resolvedDeps.joinPath(tempDir, 'response.json');

  try {
    const result = resolvedDeps.runAppCommandCaptureOutput(appPath, [
      '--update',
      '--update-launcher-path',
      scriptPath,
      '--update-response-path',
      responsePath,
    ]);
    if (result.error) {
      throw result.error;
    }
    if (result.status !== 0) {
      throw new Error(`SubMiner update command exited with status ${result.status}.`);
    }
    const response = await resolvedDeps.waitForUpdateResponse(responsePath);
    if (!response.ok) {
      throw new Error(response.error || 'SubMiner update check failed.');
    }
    return true;
  } finally {
    resolvedDeps.removeDir(tempDir);
  }
}
