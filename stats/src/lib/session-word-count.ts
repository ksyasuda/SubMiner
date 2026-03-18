type SessionWordCountLike = {
  wordsSeen: number;
  tokensSeen: number;
};

export function getSessionDisplayWordCount(value: SessionWordCountLike): number {
  return value.tokensSeen > 0 ? value.tokensSeen : value.wordsSeen;
}
