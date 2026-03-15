import { useMediaDetail } from '../../hooks/useMediaDetail';
import { MediaHeader } from './MediaHeader';
import { MediaWatchChart } from './MediaWatchChart';
import { MediaSessionList } from './MediaSessionList';

interface MediaDetailViewProps {
  videoId: number;
  onBack: () => void;
}

export function MediaDetailView({ videoId, onBack }: MediaDetailViewProps) {
  const { data, loading, error } = useMediaDetail(videoId);

  if (loading) return <div className="text-ctp-overlay2 p-4">Loading...</div>;
  if (error) return <div className="text-ctp-red p-4">Error: {error}</div>;
  if (!data?.detail) return <div className="text-ctp-overlay2 p-4">Media not found</div>;

  return (
    <div className="space-y-4">
      <button
        type="button"
        onClick={onBack}
        className="text-sm text-ctp-blue hover:text-ctp-sapphire transition-colors"
      >
        &larr; Back to Library
      </button>
      <MediaHeader detail={data.detail} />
      <MediaWatchChart rollups={data.rollups} />
      <MediaSessionList sessions={data.sessions} />
    </div>
  );
}
