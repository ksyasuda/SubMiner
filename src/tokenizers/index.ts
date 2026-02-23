import { MecabTokenizer } from '../mecab-tokenizer';
import { MecabStatus, Token } from '../types';

export interface TokenizerProvider {
  id: string;
  checkAvailability: () => Promise<boolean>;
  tokenize: (text: string) => Promise<Token[] | null>;
  getStatus: () => MecabStatus;
  setEnabled: (enabled: boolean) => void;
}

type TokenizerProviderFactory = () => TokenizerProvider;

const tokenizerProviderFactories = new Map<string, TokenizerProviderFactory>();

export function registerTokenizerProvider(id: string, factory: TokenizerProviderFactory): void {
  if (tokenizerProviderFactories.has(id)) {
    return;
  }
  tokenizerProviderFactories.set(id, factory);
}

function registerDefaultTokenizerProviders(): void {
  registerTokenizerProvider('mecab', () => {
    const mecab = new MecabTokenizer();
    return {
      id: 'mecab',
      checkAvailability: () => mecab.checkAvailability(),
      tokenize: (text: string) => mecab.tokenize(text),
      getStatus: () => mecab.getStatus(),
      setEnabled: (enabled: boolean) => mecab.setEnabled(enabled),
    };
  });
}

registerDefaultTokenizerProviders();
