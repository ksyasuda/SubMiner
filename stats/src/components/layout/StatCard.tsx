interface StatCardProps {
  label: string;
  value: string;
  subValue?: string;
  color?: string;
  trend?: { direction: 'up' | 'down' | 'flat'; text: string };
}

export function StatCard({ label, value, subValue, color = 'text-ctp-text', trend }: StatCardProps) {
  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg p-4 text-center">
      <div className={`text-2xl font-bold ${color}`}>{value}</div>
      <div className="text-xs text-ctp-subtext0 mt-1 uppercase tracking-wide">{label}</div>
      {subValue && (
        <div className="text-xs text-ctp-overlay2 mt-1">{subValue}</div>
      )}
      {trend && (
        <div className={`text-xs mt-1 ${trend.direction === 'up' ? 'text-ctp-green' : trend.direction === 'down' ? 'text-ctp-red' : 'text-ctp-overlay2'}`}>
          {trend.direction === 'up' ? '\u25B2' : trend.direction === 'down' ? '\u25BC' : '\u2014'} {trend.text}
        </div>
      )}
    </div>
  );
}
