import { formatSentenceFurigana } from '../../../anki-integration/sentence-furigana';
import { requestYomitanParseResults } from './yomitan-parser-runtime';

export async function generateSentenceFurigana(
  text: string,
  highlightedText: string | undefined,
  deps: Parameters<typeof requestYomitanParseResults>[1],
  logger: Parameters<typeof requestYomitanParseResults>[2],
): Promise<string | null> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    const results = await Promise.race([
      requestYomitanParseResults(text, deps, logger),
      new Promise<never>((_, reject) => {
        timer = setTimeout(
          () => reject(new Error('Sentence furigana generation timed out')),
          10_000,
        );
      }),
    ]);
    return formatSentenceFurigana(text, results, highlightedText);
  } catch (error) {
    logger.warn?.('Failed to generate sentence furigana:', error);
    return null;
  } finally {
    clearTimeout(timer);
  }
}
