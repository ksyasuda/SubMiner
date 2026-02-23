import axios from 'axios';

export interface TranslationRequest {
  sentence: string;
  apiKey: string;
  baseUrl: string;
  model: string;
  targetLanguage: string;
  systemPrompt: string;
  timeoutMs?: number;
}

export interface TranslationProvider {
  id: string;
  translate: (request: TranslationRequest) => Promise<string | null>;
}

type TranslationProviderFactory = () => TranslationProvider;

const translationProviderFactories = new Map<string, TranslationProviderFactory>();

export function registerTranslationProvider(id: string, factory: TranslationProviderFactory): void {
  if (translationProviderFactories.has(id)) {
    return;
  }
  translationProviderFactories.set(id, factory);
}

export function createTranslationProvider(id = 'openai-compatible'): TranslationProvider | null {
  const factory = translationProviderFactories.get(id);
  if (!factory) return null;
  return factory();
}

function extractAiText(content: unknown): string {
  if (typeof content === 'string') {
    return content.trim();
  }
  if (!Array.isArray(content)) {
    return '';
  }
  const parts: string[] = [];
  for (const item of content) {
    if (
      item &&
      typeof item === 'object' &&
      'type' in item &&
      (item as { type?: unknown }).type === 'text' &&
      'text' in item &&
      typeof (item as { text?: unknown }).text === 'string'
    ) {
      parts.push((item as { text: string }).text);
    }
  }
  return parts.join('').trim();
}

function normalizeOpenAiBaseUrl(baseUrl: string): string {
  const trimmed = baseUrl.trim().replace(/\/+$/, '');
  if (/\/v1$/i.test(trimmed)) {
    return trimmed;
  }
  return `${trimmed}/v1`;
}

function registerDefaultTranslationProviders(): void {
  registerTranslationProvider('openai-compatible', () => ({
    id: 'openai-compatible',
    translate: async (request: TranslationRequest): Promise<string | null> => {
      const response = await axios.post(
        `${normalizeOpenAiBaseUrl(request.baseUrl)}/chat/completions`,
        {
          model: request.model,
          temperature: 0,
          messages: [
            { role: 'system', content: request.systemPrompt },
            {
              role: 'user',
              content: `Translate this text to ${request.targetLanguage}:\n\n${request.sentence}`,
            },
          ],
        },
        {
          headers: {
            Authorization: `Bearer ${request.apiKey}`,
            'Content-Type': 'application/json',
          },
          timeout: request.timeoutMs ?? 15000,
        },
      );

      const content = (response.data as { choices?: unknown[] })?.choices?.[0] as
        | { message?: { content?: unknown } }
        | undefined;
      const translated = extractAiText(content?.message?.content);
      return translated || null;
    },
  }));
}

registerDefaultTranslationProviders();
