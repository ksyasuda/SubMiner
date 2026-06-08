import { useEffect, useState } from 'react';
import { appendCoverRetryToken, getCoverRetryDelayMs } from '../../lib/cover-retry';

interface RetryingCoverImageProps {
  src: string;
  alt: string;
  fallbackLabel: string;
  className?: string;
  fallbackTextClassName?: string;
  loading?: 'eager' | 'lazy';
}

export function RetryingCoverImage({
  src,
  alt,
  fallbackLabel,
  className = '',
  fallbackTextClassName = 'text-2xl',
  loading = 'lazy',
}: RetryingCoverImageProps) {
  const [failed, setFailed] = useState(false);
  const [retryToken, setRetryToken] = useState(0);
  const fallbackChar = fallbackLabel.charAt(0) || '?';

  useEffect(() => {
    setFailed(false);
    setRetryToken(0);
  }, [src]);

  useEffect(() => {
    if (!failed) return;
    const timer = setTimeout(() => {
      setRetryToken((value) => value + 1);
      setFailed(false);
    }, getCoverRetryDelayMs(retryToken));
    return () => clearTimeout(timer);
  }, [failed, retryToken]);

  if (failed) {
    return (
      <div
        className={`bg-ctp-surface2 flex items-center justify-center text-ctp-overlay2 ${fallbackTextClassName} font-bold ${className}`}
      >
        {fallbackChar}
      </div>
    );
  }

  return (
    <img
      src={appendCoverRetryToken(src, retryToken)}
      alt={alt}
      loading={loading}
      className={`object-cover bg-ctp-surface2 ${className}`}
      onError={() => setFailed(true)}
      onLoad={() => setFailed(false)}
    />
  );
}
