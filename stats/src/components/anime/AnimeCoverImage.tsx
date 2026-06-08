import { RetryingCoverImage } from '../common/RetryingCoverImage';
import { getStatsClient } from '../../hooks/useStatsApi';

interface AnimeCoverImageProps {
  animeId: number;
  title: string;
  coverRetryToken?: number;
  className?: string;
}

export function AnimeCoverImage({
  animeId,
  title,
  coverRetryToken = 0,
  className = '',
}: AnimeCoverImageProps) {
  const src = getStatsClient().getAnimeCoverUrl(animeId, coverRetryToken);

  return <RetryingCoverImage src={src} alt={title} fallbackLabel={title} className={className} />;
}
