import { TokenizerProvider } from '../../tokenizers';
import { Token } from '../../types';

export async function tokenizeStage(
  tokenizerProvider: TokenizerProvider | null,
  input: string,
): Promise<Token[] | null> {
  if (!tokenizerProvider || !input) {
    return null;
  }
  return tokenizerProvider.tokenize(input);
}
