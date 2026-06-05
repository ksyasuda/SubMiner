import { RetryingCoverImage } from '../common/RetryingCoverImage';
import { getStatsClient } from '../../hooks/useStatsApi';

interface AnimeCoverImageProps {
  animeId: number;
  title: string;
  className?: string;
}

export function AnimeCoverImage({ animeId, title, className = '' }: AnimeCoverImageProps) {
  const src = getStatsClient().getAnimeCoverUrl(animeId);

  return <RetryingCoverImage src={src} alt={title} fallbackLabel={title} className={className} />;
}
