import { useId, useState } from 'react';

interface CollapsibleSectionProps {
  title: string;
  defaultOpen?: boolean;
  children: React.ReactNode;
}

export function CollapsibleSection({
  title,
  defaultOpen = true,
  children,
}: CollapsibleSectionProps) {
  const [open, setOpen] = useState(defaultOpen);
  const contentId = useId();

  return (
    <div className="bg-ctp-surface0 border border-ctp-surface1 rounded-lg">
      <button
        type="button"
        onClick={() => setOpen(!open)}
        aria-expanded={open}
        aria-controls={contentId}
        className="w-full flex items-center justify-between p-4 text-left"
      >
        <h3 className="text-sm font-semibold text-ctp-text">{title}</h3>
        <span className="text-ctp-overlay2 text-xs" aria-hidden="true">
          {open ? '▲' : '▼'}
        </span>
      </button>
      {open && (
        <div id={contentId} className="px-4 pb-4">
          {children}
        </div>
      )}
    </div>
  );
}
