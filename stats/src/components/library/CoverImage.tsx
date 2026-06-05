import { RetryingCoverImage } from '../common/RetryingCoverImage';
import { resolveMediaCoverApiUrl } from '../../lib/media-library-grouping';

interface CoverImageProps {
  videoId: number;
  title: string;
  src?: string | null;
  className?: string;
}

export function CoverImage({ videoId, title, src = null, className = '' }: CoverImageProps) {
  const resolvedSrc = src?.trim() || resolveMediaCoverApiUrl(videoId);

  return (
    <RetryingCoverImage src={resolvedSrc} alt={title} fallbackLabel={title} className={className} />
  );
}
