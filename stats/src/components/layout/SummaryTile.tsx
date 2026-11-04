interface SummaryTileProps {
  label: string;
  value: string;
  detail: string | null;
  /** Smaller value for text such as dates. */
  compact?: boolean;
}

/** One headline number. Group several in a grid for a summary strip. */
export function SummaryTile({ label, value, detail, compact = false }: SummaryTileProps) {
  return (
    <div className="min-w-0">
      <div className="text-[11px] text-ctp-subtext0">{label}</div>
      <div
        className={`font-semibold leading-tight text-ctp-text ${compact ? 'mt-1 text-sm' : 'text-2xl'}`}
      >
        {value}
      </div>
      {detail ? <div className="text-[11px] text-ctp-overlay2">{detail}</div> : null}
    </div>
  );
}
