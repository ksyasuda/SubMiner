import axios from "axios";

import { DEFAULT_ANKI_CONNECT_CONFIG } from "../config";

const DEFAULT_AI_SYSTEM_PROMPT =
  "You are a translation engine. Return only the translated text with no explanations.";

export function extractAiText(content: unknown): string {
  if (typeof content === "string") {
    return content.trim();
  }
  if (!Array.isArray(content)) {
    return "";
  }

  const parts: string[] = [];
  for (const item of content) {
    if (
      item &&
      typeof item === "object" &&
      "type" in item &&
      (item as { type?: unknown }).type === "text" &&
      "text" in item &&
      typeof (item as { text?: unknown }).text === "string"
    ) {
      parts.push((item as { text: string }).text);
    }
  }

  return parts.join("").trim();
}

export function normalizeOpenAiBaseUrl(baseUrl: string): string {
  const trimmed = baseUrl.trim().replace(/\/+$/, "");
  if (/\/v1$/i.test(trimmed)) {
    return trimmed;
  }
  return `${trimmed}/v1`;
}

export interface AiTranslateRequest {
  sentence: string;
  apiKey: string;
  baseUrl?: string;
  model?: string;
  targetLanguage?: string;
  systemPrompt?: string;
}

export interface AiTranslateCallbacks {
  logWarning: (message: string) => void;
}

export interface AiSentenceTranslationInput {
  sentence: string;
  secondarySubText?: string;
  config: {
    apiKey?: string;
    baseUrl?: string;
    model?: string;
    targetLanguage?: string;
    systemPrompt?: string;
    enabled?: boolean;
    alwaysUseAiTranslation?: boolean;
  };
}

export interface AiSentenceTranslationCallbacks {
  logWarning: (message: string) => void;
}

export async function translateSentenceWithAi(
  request: AiTranslateRequest,
  callbacks: AiTranslateCallbacks,
): Promise<string | null> {
  const aiConfig = DEFAULT_ANKI_CONNECT_CONFIG.ai;
  if (!request.apiKey.trim()) {
    return null;
  }

  const baseUrl = normalizeOpenAiBaseUrl(
    request.baseUrl || aiConfig.baseUrl || "https://openrouter.ai/api",
  );
  const model = request.model || "openai/gpt-4o-mini";
  const targetLanguage = request.targetLanguage || "English";
  const prompt = request.systemPrompt?.trim() || DEFAULT_AI_SYSTEM_PROMPT;

  try {
    const response = await axios.post(
      `${baseUrl}/chat/completions`,
      {
        model,
        temperature: 0,
        messages: [
          { role: "system", content: prompt },
          {
            role: "user",
            content: `Translate this text to ${targetLanguage}:\n\n${request.sentence}`,
          },
        ],
      },
      {
        headers: {
          Authorization: `Bearer ${request.apiKey}`,
          "Content-Type": "application/json",
        },
        timeout: 15000,
      },
    );
    const content = (response.data as { choices?: unknown[] })?.choices?.[0] as
      | { message?: { content?: unknown } }
      | undefined;
    return extractAiText(content?.message?.content) || null;
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Unknown translation error";
    callbacks.logWarning(`AI translation failed: ${message}`);
    return null;
  }
}

export async function resolveSentenceBackText(
  input: AiSentenceTranslationInput,
  callbacks: AiSentenceTranslationCallbacks,
): Promise<string> {
  const hasSecondarySub = Boolean(input.secondarySubText?.trim());
  let backText = input.secondarySubText?.trim() || "";

  const aiConfig = {
    ...DEFAULT_ANKI_CONNECT_CONFIG.ai,
    ...input.config,
  };
  const shouldAttemptAiTranslation =
    aiConfig.enabled === true &&
    (aiConfig.alwaysUseAiTranslation === true || !hasSecondarySub);

  if (!shouldAttemptAiTranslation) return backText;

  const request: AiTranslateRequest = {
    sentence: input.sentence,
    apiKey: aiConfig.apiKey ?? "",
    baseUrl: aiConfig.baseUrl,
    model: aiConfig.model,
    targetLanguage: aiConfig.targetLanguage,
    systemPrompt: aiConfig.systemPrompt,
  };

  const translated = await translateSentenceWithAi(request, {
    logWarning: (message) => callbacks.logWarning(message),
  });

  if (translated) {
    return translated;
  }

  return hasSecondarySub ? backText : input.sentence;
}
