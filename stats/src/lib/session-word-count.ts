type SessionWordCountLike = {
  tokensSeen: number;
};

export function getSessionDisplayWordCount(value: SessionWordCountLike): number {
  return value.tokensSeen;
}
