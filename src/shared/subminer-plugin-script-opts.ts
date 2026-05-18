import type { MpvBackend } from '../types/config';

export interface SubminerPluginRuntimeScriptOptConfig {
  socketPath: string;
  binaryPath?: string;
  backend: MpvBackend;
  autoStart: boolean;
  autoStartVisibleOverlay: boolean;
  autoStartPauseUntilReady: boolean;
  texthookerEnabled: boolean;
  aniskipEnabled: boolean;
  aniskipButtonKey: string;
}

function boolScriptOpt(value: boolean): 'yes' | 'no' {
  return value ? 'yes' : 'no';
}

export function buildSubminerPluginRuntimeScriptOptParts(
  runtimeConfig: SubminerPluginRuntimeScriptOptConfig,
  fallbackAppPath: string,
): string[] {
  const binaryPath = runtimeConfig.binaryPath?.trim() || fallbackAppPath;
  return [
    `subminer-binary_path=${binaryPath}`,
    `subminer-socket_path=${runtimeConfig.socketPath}`,
    `subminer-backend=${runtimeConfig.backend}`,
    `subminer-auto_start=${boolScriptOpt(runtimeConfig.autoStart)}`,
    `subminer-auto_start_visible_overlay=${boolScriptOpt(runtimeConfig.autoStartVisibleOverlay)}`,
    `subminer-auto_start_pause_until_ready=${boolScriptOpt(
      runtimeConfig.autoStartPauseUntilReady,
    )}`,
    `subminer-texthooker_enabled=${boolScriptOpt(runtimeConfig.texthookerEnabled)}`,
    `subminer-aniskip_enabled=${boolScriptOpt(runtimeConfig.aniskipEnabled)}`,
    `subminer-aniskip_button_key=${runtimeConfig.aniskipButtonKey}`,
  ];
}
