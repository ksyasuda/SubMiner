import type { AiConfig } from '../types/integrations';
import { requestAiChatCompletion } from '../ai/client';

const DEFAULT_AI_SYSTEM_PROMPT =
  'You are a translation engine. Return only the translated text with no explanations.';

export interface AiTranslateRequest {
  sentence: string;
  apiKey: string;
  baseUrl?: string;
  model?: string;
  systemPrompt?: string;
  requestTimeoutMs?: number;
}

export interface AiTranslateCallbacks {
  logWarning: (message: string) => void;
}

export interface AiSentenceTranslationInput {
  sentence: string;
  secondarySubText?: string;
  aiEnabled: boolean;
  aiConfig: AiConfig;
}

export interface AiSentenceTranslationCallbacks {
  logWarning: (message: string) => void;
  translateSentence?: (
    request: AiTranslateRequest,
    callbacks: AiTranslateCallbacks,
  ) => Promise<string | null>;
}

export async function translateSentenceWithAi(
  request: AiTranslateRequest,
  callbacks: AiTranslateCallbacks,
): Promise<string | null> {
  if (!request.apiKey.trim()) {
    return null;
  }
  const prompt = request.systemPrompt?.trim() || DEFAULT_AI_SYSTEM_PROMPT;

  return requestAiChatCompletion(
    {
      apiKey: request.apiKey,
      baseUrl: request.baseUrl,
      model: request.model,
      timeoutMs: request.requestTimeoutMs,
      messages: [
        { role: 'system', content: prompt },
        {
          role: 'user',
          content: `Translate this text to English:\n\n${request.sentence}`,
        },
      ],
    },
    {
      logWarning: (message) =>
        callbacks.logWarning(message.replace(/^AI request failed:/, 'AI translation failed:')),
    },
  );
}

export async function resolveSentenceBackText(
  input: AiSentenceTranslationInput,
  callbacks: AiSentenceTranslationCallbacks,
): Promise<string> {
  const hasSecondarySub = Boolean(input.secondarySubText?.trim());
  let backText = input.secondarySubText?.trim() || '';

  const shouldAttemptAiTranslation =
    input.aiEnabled === true && input.aiConfig.enabled === true && !hasSecondarySub;

  if (!shouldAttemptAiTranslation) return backText;

  const translateSentence = callbacks.translateSentence ?? translateSentenceWithAi;

  const request: AiTranslateRequest = {
    sentence: input.sentence,
    apiKey: input.aiConfig.apiKey ?? '',
    baseUrl: input.aiConfig.baseUrl,
    model: input.aiConfig.model,
    systemPrompt: input.aiConfig.systemPrompt,
    requestTimeoutMs: input.aiConfig.requestTimeoutMs,
  };

  const translated = await translateSentence(request, {
    logWarning: (message) => callbacks.logWarning(message),
  });

  if (translated) {
    return translated;
  }

  return hasSecondarySub ? backText : input.sentence;
}
