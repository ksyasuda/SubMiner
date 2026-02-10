import { mergeTokens as defaultMergeTokens } from "../token-merger";
import { MergedToken, Token } from "../types";

export interface TokenMergerProvider {
  id: string;
  merge: (tokens: Token[]) => MergedToken[];
}

type TokenMergerProviderFactory = () => TokenMergerProvider;

const tokenMergerProviderFactories = new Map<string, TokenMergerProviderFactory>();

export function registerTokenMergerProvider(
  id: string,
  factory: TokenMergerProviderFactory,
): void {
  if (tokenMergerProviderFactories.has(id)) {
    return;
  }
  tokenMergerProviderFactories.set(id, factory);
}

export function getRegisteredTokenMergerProviderIds(): string[] {
  return Array.from(tokenMergerProviderFactories.keys());
}

export function createTokenMergerProvider(
  id = "default",
): TokenMergerProvider | null {
  const factory = tokenMergerProviderFactories.get(id);
  if (!factory) return null;
  return factory();
}

function registerDefaultTokenMergerProviders(): void {
  registerTokenMergerProvider("default", () => ({
    id: "default",
    merge: (tokens: Token[]) => defaultMergeTokens(tokens),
  }));
}

registerDefaultTokenMergerProviders();
