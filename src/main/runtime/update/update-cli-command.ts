import fs from 'node:fs';
import path from 'node:path';
import type { CliArgs, CliCommandSource } from '../../../cli/args';
import type { UpdateCheckRequest, UpdateCheckResult } from './update-service';

export interface UpdateCliCommandResponse {
  ok: boolean;
  status?: UpdateCheckResult['status'];
  version?: string;
  error?: string;
}

export interface UpdateCliCommandDeps {
  checkForUpdates: (request: UpdateCheckRequest) => Promise<UpdateCheckResult>;
  writeResponse: (responsePath: string, payload: UpdateCliCommandResponse) => void;
  logWarn: (message: string, error?: unknown) => void;
}

export function writeUpdateCliCommandResponse(
  responsePath: string,
  payload: UpdateCliCommandResponse,
): void {
  fs.mkdirSync(path.dirname(responsePath), { recursive: true });
  fs.writeFileSync(responsePath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
}

function responseFromResult(result: UpdateCheckResult): UpdateCliCommandResponse {
  const response: UpdateCliCommandResponse = {
    ok: result.status !== 'failed',
    status: result.status,
  };
  if (result.version !== undefined) response.version = result.version;
  if (result.error !== undefined) response.error = result.error;
  return response;
}

function writeResponseSafe(
  responsePath: string | undefined,
  payload: UpdateCliCommandResponse,
  deps: Pick<UpdateCliCommandDeps, 'writeResponse' | 'logWarn'>,
): void {
  if (!responsePath) return;
  try {
    deps.writeResponse(responsePath, payload);
  } catch (error) {
    deps.logWarn(`Failed to write update response: ${responsePath}`, error);
  }
}

export async function runUpdateCliCommand(
  args: Pick<CliArgs, 'updateLauncherPath' | 'updateResponsePath'>,
  _source: CliCommandSource,
  deps: UpdateCliCommandDeps,
): Promise<UpdateCheckResult> {
  try {
    const result = await deps.checkForUpdates({
      source: args.updateLauncherPath ? 'launcher' : 'manual',
      launcherPath: args.updateLauncherPath,
    });
    writeResponseSafe(args.updateResponsePath, responseFromResult(result), deps);
    return result;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    writeResponseSafe(
      args.updateResponsePath,
      { ok: false, status: 'failed', error: message },
      deps,
    );
    throw error;
  }
}
