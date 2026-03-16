interface StatCardProps {
  label: string;
  value: string;
  subValue?: string;
  color?: string;
  trend?: { direction: 'up' | 'down' | 'flat'; text: string };
}

const COLOR_TO_BORDER: Record<string, string> = {
  'text-ctp-blue': 'border-l-ctp-blue',
  'text-ctp-green': 'border-l-ctp-green',
  'text-ctp-mauve': 'border-l-ctp-mauve',
  'text-ctp-peach': 'border-l-ctp-peach',
  'text-ctp-teal': 'border-l-ctp-teal',
  'text-ctp-lavender': 'border-l-ctp-lavender',
  'text-ctp-red': 'border-l-ctp-red',
  'text-ctp-yellow': 'border-l-ctp-yellow',
  'text-ctp-sapphire': 'border-l-ctp-sapphire',
  'text-ctp-sky': 'border-l-ctp-sky',
  'text-ctp-flamingo': 'border-l-ctp-flamingo',
  'text-ctp-maroon': 'border-l-ctp-maroon',
  'text-ctp-pink': 'border-l-ctp-pink',
  'text-ctp-text': 'border-l-ctp-surface2',
};

export function StatCard({
  label,
  value,
  subValue,
  color = 'text-ctp-text',
  trend,
}: StatCardProps) {
  const borderClass = COLOR_TO_BORDER[color] ?? 'border-l-ctp-surface2';

  return (
    <div
      className={`bg-ctp-surface0 border border-ctp-surface1 border-l-[3px] ${borderClass} rounded-lg p-4 text-center`}
    >
      <div className={`text-2xl font-bold font-mono tabular-nums ${color}`}>{value}</div>
      <div className="text-xs text-ctp-subtext0 mt-1 uppercase tracking-wide">{label}</div>
      {subValue && <div className="text-xs text-ctp-overlay2 mt-1">{subValue}</div>}
      {trend && (
        <div
          className={`text-xs mt-1 font-mono tabular-nums ${trend.direction === 'up' ? 'text-ctp-green' : trend.direction === 'down' ? 'text-ctp-red' : 'text-ctp-overlay2'}`}
        >
          {trend.direction === 'up' ? '\u25B2' : trend.direction === 'down' ? '\u25BC' : '\u2014'}{' '}
          {trend.text}
        </div>
      )}
    </div>
  );
}
