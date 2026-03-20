interface TooltipProps {
  text: string;
  children: React.ReactNode;
}

export function Tooltip({ text, children }: TooltipProps) {
  return (
    <div className="group/tip relative">
      {children}
      <div
        role="tooltip"
        className="pointer-events-none absolute bottom-full left-1/2 -translate-x-1/2 mb-2 z-50
          max-w-56 px-2.5 py-1.5 rounded-md text-xs text-ctp-text bg-ctp-surface2 border border-ctp-overlay0 shadow-lg
          opacity-0 scale-95 transition-all duration-150
          group-hover/tip:opacity-100 group-hover/tip:scale-100"
      >
        {text}
        <div className="absolute top-full left-1/2 -translate-x-1/2 -mt-px border-4 border-transparent border-t-ctp-surface2" />
      </div>
    </div>
  );
}
