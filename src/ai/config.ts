import type { AiConfig, AiFeatureConfig } from '../types';

function trimToOverride(value: string | undefined): string | undefined {
  if (typeof value !== 'string') return undefined;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

export function mergeAiConfig(
  sharedConfig: AiConfig | undefined,
  featureConfig?: AiFeatureConfig | boolean | null,
): AiConfig {
  const overrides =
    featureConfig && typeof featureConfig === 'object' ? featureConfig : undefined;
  const modelOverride = trimToOverride(overrides?.model);
  const systemPromptOverride = trimToOverride(overrides?.systemPrompt);

  return {
    enabled: sharedConfig?.enabled,
    apiKey: sharedConfig?.apiKey,
    apiKeyCommand: sharedConfig?.apiKeyCommand,
    baseUrl: sharedConfig?.baseUrl,
    model: modelOverride ?? sharedConfig?.model,
    systemPrompt: systemPromptOverride ?? sharedConfig?.systemPrompt,
    requestTimeoutMs: sharedConfig?.requestTimeoutMs,
  };
}
