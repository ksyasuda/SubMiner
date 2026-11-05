interface TitleFilterProps {
  titles: string[];
  query: string;
  selectedTitle: string | null;
  onQueryChange: (query: string) => void;
  onSelectTitle: (title: string | null) => void;
}

const TITLE_OPTIONS_ID = 'trends-title-options';

/**
 * Whether an input event came from picking a datalist option rather than typing.
 * Firefox reports insertReplacementText. Chromium, which the stats window runs on,
 * sends a plain Event with no inputType, while typing always reports one.
 */
export function isDatalistSelection(event: Event): boolean {
  const inputType = 'inputType' in event ? event.inputType : undefined;
  return inputType === undefined || inputType === 'insertReplacementText';
}

/**
 * Typing narrows the title lists below. Picking a suggestion, or pressing Enter on
 * an exact title, filters the whole page to it, as choosing it from a list row does.
 * Typing alone never selects, because one title can be the start of another.
 */
export function TitleFilter({
  titles,
  query,
  selectedTitle,
  onQueryChange,
  onSelectTitle,
}: TitleFilterProps) {
  const selectIfTitle = (value: string): boolean => {
    if (!titles.includes(value)) return false;
    onSelectTitle(value);
    onQueryChange('');
    return true;
  };

  return (
    <div className="flex min-w-0 flex-wrap items-center gap-2 text-xs">
      <input
        type="search"
        aria-label="Filter by title"
        placeholder="Filter by title..."
        list={TITLE_OPTIONS_ID}
        value={query}
        onChange={(event) => {
          const value = event.target.value;
          if (!isDatalistSelection(event.nativeEvent) || !selectIfTitle(value)) {
            onQueryChange(value);
          }
        }}
        onKeyDown={(event) => {
          // Enter during IME composition confirms the conversion, not the filter.
          if (event.key === 'Enter' && !event.nativeEvent.isComposing) {
            selectIfTitle(event.currentTarget.value);
          }
        }}
        className="w-56 max-w-full rounded-lg border border-ctp-surface1 bg-ctp-surface0 px-2.5 py-1 text-xs text-ctp-text placeholder:text-ctp-overlay2 focus:border-ctp-blue focus:outline-none"
      />
      <datalist id={TITLE_OPTIONS_ID}>
        {titles.map((title) => (
          <option key={title} value={title} />
        ))}
      </datalist>
      {selectedTitle ? (
        <button
          type="button"
          aria-label={`Clear title filter: ${selectedTitle}`}
          title="Clear title filter"
          className="flex max-w-xs items-center gap-1.5 rounded-full border border-ctp-blue/60 bg-ctp-blue/12 px-2.5 py-1 text-ctp-blue transition hover:border-ctp-blue"
          onClick={() => onSelectTitle(null)}
        >
          <span className="truncate">{selectedTitle}</span>
          <span aria-hidden="true">×</span>
        </button>
      ) : null}
    </div>
  );
}
