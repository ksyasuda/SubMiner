import { useEffect, useState } from 'react';
import { getCoverImageSrc, type CoverImageMap } from '../../lib/cover-images';

interface CoverThumbnailProps {
  animeId: number | null;
  videoId: number | null;
  title: string;
  coverImages: CoverImageMap;
}

export function CoverThumbnail({ animeId, videoId, title, coverImages }: CoverThumbnailProps) {
  const [failed, setFailed] = useState(false);
  const fallbackChar = title.charAt(0) || '?';
  const fallback = (
    <div className="w-12 h-16 rounded bg-ctp-surface2 flex items-center justify-center text-ctp-overlay2 text-lg font-bold shrink-0">
      {fallbackChar}
    </div>
  );

  const src =
    animeId != null
      ? getCoverImageSrc(coverImages, 'anime', animeId)
      : getCoverImageSrc(coverImages, 'media', videoId);

  useEffect(() => {
    setFailed(false);
  }, [src]);

  if (!src || failed) {
    return fallback;
  }

  return (
    <img
      src={src}
      alt=""
      className="w-12 h-16 rounded object-cover shrink-0 bg-ctp-surface2"
      onError={() => setFailed(true)}
    />
  );
}
