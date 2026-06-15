import fs from 'node:fs';
import path from 'node:path';
import type { CliArgs, CliCommandSource } from '../../cli/args';
import {
  ensureLinuxRuntimePluginAssets,
  type EnsureLinuxRuntimePluginAssetsResult,
} from './linux-runtime-plugin-assets';

export interface EnsureLinuxRuntimePluginAssetsResponse {
  ok: boolean;
  status: 'installed' | 'already-present' | 'failed';
  path?: string;
  error?: string;
}

export interface EnsureLinuxRuntimePluginAssetsCliCommandDeps {
  ensureLinuxRuntimePluginAssets: () => Promise<EnsureLinuxRuntimePluginAssetsResult>;
  writeResponse: (responsePath: string, payload: EnsureLinuxRuntimePluginAssetsResponse) => void;
  logWarn: (message: string, error?: unknown) => void;
}

export function writeEnsureLinuxRuntimePluginAssetsCliCommandResponse(
  responsePath: string,
  payload: EnsureLinuxRuntimePluginAssetsResponse,
): void {
  fs.mkdirSync(path.dirname(responsePath), { recursive: true });
  fs.writeFileSync(responsePath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
}

function writeResponseSafe(
  responsePath: string | undefined,
  payload: EnsureLinuxRuntimePluginAssetsResponse,
  deps: Pick<EnsureLinuxRuntimePluginAssetsCliCommandDeps, 'writeResponse' | 'logWarn'>,
): void {
  if (!responsePath) return;
  try {
    deps.writeResponse(responsePath, payload);
  } catch (error) {
    deps.logWarn(`Failed to write Linux runtime plugin asset response: ${responsePath}`, error);
  }
}

const defaultDeps: EnsureLinuxRuntimePluginAssetsCliCommandDeps = {
  ensureLinuxRuntimePluginAssets: () => ensureLinuxRuntimePluginAssets(),
  writeResponse: (responsePath, payload) =>
    writeEnsureLinuxRuntimePluginAssetsCliCommandResponse(responsePath, payload),
  logWarn: () => {},
};

export async function runEnsureLinuxRuntimePluginAssetsCliCommand(
  args: Pick<
    CliArgs,
    'ensureLinuxRuntimePluginAssets' | 'ensureLinuxRuntimePluginAssetsResponsePath'
  >,
  deps: EnsureLinuxRuntimePluginAssetsCliCommandDeps = defaultDeps,
  _source: CliCommandSource = 'initial',
): Promise<EnsureLinuxRuntimePluginAssetsResult> {
  try {
    const result = await deps.ensureLinuxRuntimePluginAssets();
    writeResponseSafe(args.ensureLinuxRuntimePluginAssetsResponsePath, result, deps);
    return result;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    writeResponseSafe(
      args.ensureLinuxRuntimePluginAssetsResponsePath,
      { ok: false, status: 'failed', error: message },
      deps,
    );
    throw error;
  }
}
