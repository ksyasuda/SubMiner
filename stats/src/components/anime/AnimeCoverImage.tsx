import { useState } from 'react';
import { getStatsClient } from '../../hooks/useStatsApi';

interface AnimeCoverImageProps {
  animeId: number;
  title: string;
  className?: string;
}

export function AnimeCoverImage({ animeId, title, className = '' }: AnimeCoverImageProps) {
  const [failed, setFailed] = useState(false);
  const fallbackChar = title.charAt(0) || '?';

  if (failed) {
    return (
      <div
        className={`bg-ctp-surface2 flex items-center justify-center text-ctp-overlay2 text-2xl font-bold ${className}`}
      >
        {fallbackChar}
      </div>
    );
  }

  const src = getStatsClient().getAnimeCoverUrl(animeId);

  return (
    <img
      src={src}
      alt={title}
      loading="lazy"
      className={`object-cover bg-ctp-surface2 ${className}`}
      onError={() => setFailed(true)}
    />
  );
}
