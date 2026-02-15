import { TokenMergerProvider } from "../../token-mergers";
import { MergedToken, Token } from "../../types";

export function mergeStage(
  mergerProvider: TokenMergerProvider | null,
  tokens: Token[] | null,
): MergedToken[] | null {
  if (!mergerProvider || !tokens || tokens.length === 0) {
    return null;
  }
  return mergerProvider.merge(tokens);
}
