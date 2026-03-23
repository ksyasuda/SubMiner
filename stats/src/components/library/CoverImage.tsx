import { useEffect, useState } from 'react';
import { resolveMediaCoverApiUrl } from '../../lib/media-library-grouping';

interface CoverImageProps {
  videoId: number;
  title: string;
  src?: string | null;
  className?: string;
}

export function CoverImage({ videoId, title, src = null, className = '' }: CoverImageProps) {
  const [failed, setFailed] = useState(false);
  const fallbackChar = title.charAt(0) || '?';
  const resolvedSrc = src?.trim() || resolveMediaCoverApiUrl(videoId);

  useEffect(() => {
    setFailed(false);
  }, [resolvedSrc]);

  if (failed) {
    return (
      <div
        className={`bg-ctp-surface2 flex items-center justify-center text-ctp-overlay2 text-2xl font-bold ${className}`}
      >
        {fallbackChar}
      </div>
    );
  }

  return (
    <img
      src={resolvedSrc}
      alt={title}
      loading="lazy"
      className={`object-cover bg-ctp-surface2 ${className}`}
      onError={() => setFailed(true)}
    />
  );
}
