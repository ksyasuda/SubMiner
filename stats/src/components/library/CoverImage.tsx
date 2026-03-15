import { useState } from 'react';
import { BASE_URL } from '../../lib/api-client';

interface CoverImageProps {
  videoId: number;
  title: string;
  className?: string;
}

export function CoverImage({ videoId, title, className = '' }: CoverImageProps) {
  const [failed, setFailed] = useState(false);
  const fallbackChar = title.charAt(0) || '?';

  if (failed) {
    return (
      <div className={`bg-ctp-surface2 flex items-center justify-center text-ctp-overlay2 text-2xl font-bold ${className}`}>
        {fallbackChar}
      </div>
    );
  }

  return (
    <img
      src={`${BASE_URL}/api/stats/media/${videoId}/cover`}
      alt={title}
      className={`object-cover bg-ctp-surface2 ${className}`}
      onError={() => setFailed(true)}
    />
  );
}
