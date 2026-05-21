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

function sanitizeScriptOptValue(value: string): string {
  return value
    .replace(/,/g, ' ')
    .replace(/[\r\n]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function buildSubminerPluginRuntimeScriptOptParts(
  runtimeConfig: SubminerPluginRuntimeScriptOptConfig,
  fallbackAppPath: string,
): string[] {
  const binaryPath = sanitizeScriptOptValue(runtimeConfig.binaryPath?.trim() || fallbackAppPath);
  const socketPath = sanitizeScriptOptValue(runtimeConfig.socketPath);
  const backend = sanitizeScriptOptValue(runtimeConfig.backend);
  const aniskipButtonKey = sanitizeScriptOptValue(runtimeConfig.aniskipButtonKey);
  return [
    `subminer-binary_path=${binaryPath}`,
    `subminer-socket_path=${socketPath}`,
    `subminer-backend=${backend}`,
    `subminer-auto_start=${boolScriptOpt(runtimeConfig.autoStart)}`,
    `subminer-auto_start_visible_overlay=${boolScriptOpt(runtimeConfig.autoStartVisibleOverlay)}`,
    `subminer-auto_start_pause_until_ready=${boolScriptOpt(
      runtimeConfig.autoStartPauseUntilReady,
    )}`,
    `subminer-texthooker_enabled=${boolScriptOpt(runtimeConfig.texthookerEnabled)}`,
    `subminer-aniskip_enabled=${boolScriptOpt(runtimeConfig.aniskipEnabled)}`,
    `subminer-aniskip_button_key=${aniskipButtonKey}`,
  ];
}
