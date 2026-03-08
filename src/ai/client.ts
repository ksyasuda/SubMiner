import { exec as execCallback } from 'node:child_process';
import { promisify } from 'node:util';
import axios from 'axios';

import type { AiConfig } from '../types';

const DEFAULT_AI_BASE_URL = 'https://openrouter.ai/api';
const DEFAULT_AI_MODEL = 'openai/gpt-4o-mini';
const DEFAULT_AI_TIMEOUT_MS = 15_000;

const exec = promisify(execCallback);

export function extractAiText(content: unknown): string {
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

export function normalizeOpenAiBaseUrl(baseUrl: string): string {
  const trimmed = baseUrl.trim().replace(/\/+$/, '');
  if (/\/v1$/i.test(trimmed)) {
    return trimmed;
  }
  return `${trimmed}/v1`;
}

export async function resolveAiApiKey(
  config: Pick<AiConfig, 'apiKey' | 'apiKeyCommand'>,
): Promise<string | null> {
  if (config.apiKey && config.apiKey.trim()) {
    return config.apiKey.trim();
  }
  if (config.apiKeyCommand && config.apiKeyCommand.trim()) {
    try {
      const { stdout } = await exec(config.apiKeyCommand, { timeout: 10_000 });
      const key = stdout.trim();
      return key.length > 0 ? key : null;
    } catch {
      return null;
    }
  }
  return null;
}

export interface AiChatMessage {
  role: 'system' | 'user' | 'assistant';
  content: string;
}

export interface AiChatCompletionRequest {
  apiKey: string;
  baseUrl?: string;
  model?: string;
  timeoutMs?: number;
  messages: AiChatMessage[];
}

export interface AiChatCompletionCallbacks {
  logWarning: (message: string) => void;
}

export async function requestAiChatCompletion(
  request: AiChatCompletionRequest,
  callbacks: AiChatCompletionCallbacks,
): Promise<string | null> {
  if (!request.apiKey.trim()) {
    return null;
  }

  const baseUrl = normalizeOpenAiBaseUrl(request.baseUrl || DEFAULT_AI_BASE_URL);
  const model = request.model || DEFAULT_AI_MODEL;
  const timeoutMs = request.timeoutMs ?? DEFAULT_AI_TIMEOUT_MS;

  try {
    const response = await axios.post(
      `${baseUrl}/chat/completions`,
      {
        model,
        temperature: 0,
        messages: request.messages,
      },
      {
        headers: {
          Authorization: `Bearer ${request.apiKey}`,
          'Content-Type': 'application/json',
        },
        timeout: timeoutMs,
      },
    );
    const content = (response.data as { choices?: unknown[] })?.choices?.[0] as
      | { message?: { content?: unknown } }
      | undefined;
    return extractAiText(content?.message?.content) || null;
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown AI request error';
    callbacks.logWarning(`AI request failed: ${message}`);
    return null;
  }
}
