import { TokenMergerProvider } from "../token-mergers";
import { TokenizerProvider } from "../tokenizers";
import { SubtitleData } from "../types";
import {
  normalizeDisplayText,
  normalizeTokenizerInput,
} from "./stages/normalize";
import { tokenizeStage } from "./stages/tokenize";
import { mergeStage } from "./stages/merge";

export interface SubtitlePipelineDeps {
  getTokenizer: () => TokenizerProvider | null;
  getTokenMerger: () => TokenMergerProvider | null;
}

export class SubtitlePipeline {
  private readonly deps: SubtitlePipelineDeps;

  constructor(deps: SubtitlePipelineDeps) {
    this.deps = deps;
  }

  async process(text: string): Promise<SubtitleData> {
    if (!text) {
      return { text, tokens: null };
    }

    const displayText = normalizeDisplayText(text);
    if (!displayText) {
      return { text, tokens: null };
    }

    const tokenizeText = normalizeTokenizerInput(displayText);

    try {
      const tokens = await tokenizeStage(this.deps.getTokenizer(), tokenizeText);
      const mergedTokens = mergeStage(this.deps.getTokenMerger(), tokens);
      if (!mergedTokens || mergedTokens.length === 0) {
        return { text: displayText, tokens: null };
      }
      return { text: displayText, tokens: mergedTokens };
    } catch {
      return { text: displayText, tokens: null };
    }
  }
}
