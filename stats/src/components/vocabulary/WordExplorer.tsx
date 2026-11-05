import { Fragment, useEffect, useMemo, useState } from 'react';
import { useDictionaryWords } from '../../hooks/useDictionaryWords';
import {
  epochDayToDate,
  epochMsFromDbTimestamp,
  formatNumber,
  localDayFromMs,
} from '../../lib/formatters';
import {
  readBooleanPreference,
  writeBooleanPreference,
  writeStringPreference,
} from '../../lib/preference-storage';
import { fullReading } from '../../lib/reading-utils';
import type {
  DictionaryWordRow,
  StatsDictionaryWordPos,
  StatsDictionaryWordSort,
} from '../../types/stats';
import {
  HIDE_KANA_KEY,
  HIDE_KNOWN_KEY,
  WORD_PAGE_SIZE,
  WORD_VIEW_KEY,
  WORD_VIEWS,
  buildWordQuery,
  defaultSortForView,
  loadWordView,
  nextSort,
  type WordSort,
  type WordView,
} from './word-explorer-query';

interface WordExplorerProps {
  hasKnownData: boolean;
  hideNames: boolean;
  onHideNamesChange: (hide: boolean) => void;
  /** Changes when server-side word data may have changed. */
  refreshKey: number;
  onOpenWord: (wordId: number) => void;
  onHideWord: (word: DictionaryWordRow) => void;
}

const VIEW_LABELS: Record<WordView, string> = {
  worth: 'Worth mining',
  all: 'All words',
  titles: 'Across titles',
  newest: 'Newest',
};

const POS_FILTERS: Array<{ value: StatsDictionaryWordPos; label: string }> = [
  { value: 'noun', label: 'Nouns' },
  { value: 'verb', label: 'Verbs' },
  { value: 'adjective', label: 'Adjectives' },
];

const POS_LABELS: Record<string, string> = {
  i_adjective: 'adjective',
  na_adjective: 'adjective',
  auxiliary_verb: 'auxiliary',
  bound_auxiliary: 'auxiliary',
};

const SEARCH_DEBOUNCE_MS = 250;

const COLUMNS: Array<{ id: StatsDictionaryWordSort; label: string; align: 'left' | 'right' }> = [
  { id: 'headword', label: 'Word', align: 'left' },
  { id: 'rank', label: 'Rank', align: 'right' },
  { id: 'seen', label: 'Seen', align: 'right' },
  { id: 'titles', label: 'Titles', align: 'right' },
  { id: 'firstSeen', label: 'First seen', align: 'right' },
];

function ToggleChip({
  label,
  pressed,
  onToggle,
}: {
  label: string;
  pressed: boolean;
  onToggle: () => void;
}) {
  return (
    <button
      type="button"
      aria-pressed={pressed}
      onClick={onToggle}
      className={`rounded-full border px-2.5 py-0.5 text-[11px] transition-colors ${
        pressed
          ? 'border-ctp-blue/70 text-ctp-text'
          : 'border-ctp-surface1 text-ctp-overlay2 hover:text-ctp-subtext0'
      }`}
    >
      {label}
    </button>
  );
}

function formatFirstSeen(timestamp: number): string {
  const day = localDayFromMs(epochMsFromDbTimestamp(timestamp));
  return epochDayToDate(day).toLocaleDateString(undefined, {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
  });
}

function rowReading(row: DictionaryWordRow): string {
  const reading = fullReading(row.headword, row.reading);
  // `fullReading` normalizes katakana, so カレー must not render its own hiragana.
  return reading === fullReading(row.headword, row.headword) ? '' : reading;
}

function describeResults(view: WordView, searching: boolean, hasKnownData: boolean): string {
  if (searching) return 'matching words';
  if (view === 'worth') return hasKnownData ? 'common words not yet mined' : 'common words';
  if (view === 'titles') return 'words across titles';
  return 'words';
}

/** The one word list: server-filtered, so it reaches every word you have seen. */
export function WordExplorer({
  hasKnownData,
  hideNames,
  onHideNamesChange,
  refreshKey,
  onOpenWord,
  onHideWord,
}: WordExplorerProps) {
  const [view, setView] = useState<WordView>(() => loadWordView());
  const [sort, setSort] = useState<WordSort>(() => defaultSortForView(loadWordView()));
  const [searchDraft, setSearchDraft] = useState('');
  const [search, setSearch] = useState('');
  const [hideKnown, setHideKnown] = useState(() => readBooleanPreference(HIDE_KNOWN_KEY, true));
  const [hideKana, setHideKana] = useState(() => readBooleanPreference(HIDE_KANA_KEY, true));
  const [partOfSpeech, setPartOfSpeech] = useState<StatsDictionaryWordPos | null>(null);
  const [page, setPage] = useState(0);
  const [expandedWord, setExpandedWord] = useState<string | null>(null);

  useEffect(() => {
    const timer = setTimeout(() => {
      setSearch(searchDraft);
      setPage(0);
    }, SEARCH_DEBOUNCE_MS);
    return () => clearTimeout(timer);
  }, [searchDraft]);

  const query = useMemo(
    () =>
      buildWordQuery({
        view,
        search,
        hideKnown: hasKnownData && hideKnown,
        hideNames,
        hideKana,
        partOfSpeech,
        sort,
        page,
      }),
    [view, search, hasKnownData, hideKnown, hideNames, hideKana, partOfSpeech, sort, page],
  );
  const { page: results, loading, error } = useDictionaryWords(query, refreshKey);

  const selectView = (next: WordView) => {
    setView(next);
    setSort(defaultSortForView(next));
    setPage(0);
    writeStringPreference(WORD_VIEW_KEY, next);
  };
  const applyFilter = (update: () => void) => {
    update();
    setPage(0);
  };

  const total = results?.total ?? 0;
  const rows = results?.rows ?? [];
  const firstRow = total === 0 ? 0 : page * WORD_PAGE_SIZE + 1;
  const lastRow = Math.min((page + 1) * WORD_PAGE_SIZE, total);
  const searching = search.trim() !== '';

  return (
    <section
      aria-label="Words"
      className="rounded-lg border border-ctp-surface1 bg-ctp-surface0 p-4"
    >
      <div className="flex flex-wrap items-center gap-2">
        <div
          role="group"
          aria-label="Word list view"
          className="flex rounded-lg border border-ctp-surface1 p-0.5"
        >
          {WORD_VIEWS.map((option) => (
            <button
              key={option}
              type="button"
              aria-pressed={view === option}
              onClick={() => selectView(option)}
              className={`rounded-md px-2.5 py-1 text-xs transition-colors ${
                view === option
                  ? 'bg-ctp-surface2 text-ctp-text'
                  : 'text-ctp-overlay2 hover:text-ctp-subtext0'
              }`}
            >
              {VIEW_LABELS[option]}
            </button>
          ))}
        </div>
        <input
          type="search"
          aria-label="Search all words"
          placeholder="Search all words..."
          value={searchDraft}
          onChange={(event) => setSearchDraft(event.target.value)}
          className="min-w-40 flex-1 rounded-lg border border-ctp-surface1 bg-transparent px-2.5 py-1 text-xs text-ctp-text placeholder:text-ctp-overlay2 focus:border-ctp-blue focus:outline-none"
        />
      </div>

      {searching ? null : (
        <div className="mt-2 flex flex-wrap items-center gap-1.5">
          {hasKnownData && view !== 'worth' ? (
            <ToggleChip
              label="Hide known"
              pressed={hideKnown}
              onToggle={() =>
                applyFilter(() => {
                  setHideKnown(!hideKnown);
                  writeBooleanPreference(HIDE_KNOWN_KEY, !hideKnown);
                })
              }
            />
          ) : null}
          <ToggleChip
            label="Hide names"
            pressed={hideNames}
            onToggle={() => applyFilter(() => onHideNamesChange(!hideNames))}
          />
          <ToggleChip
            label="Hide kana"
            pressed={hideKana}
            onToggle={() =>
              applyFilter(() => {
                setHideKana(!hideKana);
                writeBooleanPreference(HIDE_KANA_KEY, !hideKana);
              })
            }
          />
          <span aria-hidden="true" className="mx-1 h-3 w-px bg-ctp-surface1" />
          {POS_FILTERS.map(({ value, label }) => (
            <ToggleChip
              key={value}
              label={label}
              pressed={partOfSpeech === value}
              onToggle={() =>
                applyFilter(() => setPartOfSpeech(partOfSpeech === value ? null : value))
              }
            />
          ))}
        </div>
      )}

      {error ? (
        <div role="alert" className="mt-3 text-xs text-ctp-red">
          Words failed to load: {error}
        </div>
      ) : null}

      <div
        aria-busy={loading}
        className={`mt-3 overflow-x-auto transition-opacity ${loading && results ? 'opacity-60' : ''}`}
      >
        <table className="w-full text-xs">
          <thead>
            <tr className="border-b border-ctp-surface1 text-ctp-overlay2">
              {COLUMNS.map((column) => {
                const active = sort.sort === column.id;
                return (
                  <th
                    key={column.id}
                    scope="col"
                    aria-sort={active ? (sort.descending ? 'descending' : 'ascending') : undefined}
                    className={`pb-2 pl-3 font-medium first:pl-0 ${
                      column.align === 'right' ? 'text-right' : 'text-left'
                    }`}
                  >
                    <button
                      type="button"
                      onClick={() => {
                        setSort(nextSort(sort, column.id));
                        setPage(0);
                      }}
                      className={`whitespace-nowrap transition-colors hover:text-ctp-text ${
                        active ? 'text-ctp-text' : ''
                      }`}
                    >
                      {column.label}
                      {active ? (sort.descending ? ' ▼' : ' ▲') : ''}
                    </button>
                  </th>
                );
              })}
              <th scope="col" className="pb-2 pl-3">
                <span className="sr-only">Actions</span>
              </th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => {
              const expanded = expandedWord === row.headword;
              const reading = rowReading(row);
              return (
                <Fragment key={row.headword}>
                  <tr className="border-b border-ctp-surface1/50 transition-colors last:border-0 hover:bg-ctp-surface1/30">
                    <td className="py-1.5">
                      <div className="flex min-w-0 items-baseline gap-2">
                        {row.formCount > 1 ? (
                          <button
                            type="button"
                            aria-expanded={expanded}
                            aria-label={`Show forms of ${row.headword}`}
                            onClick={() => setExpandedWord(expanded ? null : row.headword)}
                            className={`w-3 shrink-0 text-[9px] text-ctp-overlay2 transition-transform hover:text-ctp-text ${
                              expanded ? 'rotate-90' : ''
                            }`}
                          >
                            ▶
                          </button>
                        ) : (
                          <span aria-hidden="true" className="w-3 shrink-0" />
                        )}
                        <button
                          type="button"
                          lang="ja"
                          onClick={() => onOpenWord(row.wordId)}
                          className="text-left text-[15px] font-medium text-ctp-text transition-colors hover:text-ctp-blue"
                        >
                          {row.headword}
                        </button>
                        {reading ? (
                          <span lang="ja" className="text-[11px] text-ctp-overlay2">
                            {reading}
                          </span>
                        ) : null}
                        <span className="text-[11px] text-ctp-overlay1">
                          {row.partOfSpeech
                            ? (POS_LABELS[row.partOfSpeech] ?? row.partOfSpeech.replace(/_/g, ' '))
                            : ''}
                        </span>
                        {row.known ? (
                          <span className="text-[11px] text-ctp-green">Known</span>
                        ) : null}
                        {row.isName ? (
                          <span className="text-[11px] text-ctp-overlay1">Name</span>
                        ) : null}
                        {row.streamOnly ? (
                          <span
                            title="Every line with this word came from a stream, so it can't be mined yet."
                            className="text-[11px] whitespace-nowrap text-ctp-peach"
                          >
                            Stream only
                          </span>
                        ) : null}
                      </div>
                    </td>
                    <td className="py-1.5 pl-3 text-right tabular-nums text-ctp-subtext0">
                      {row.rank === null ? '—' : formatNumber(row.rank)}
                    </td>
                    <td className="py-1.5 pl-3 text-right tabular-nums text-ctp-text">
                      <div className="flex items-center justify-end gap-2">
                        {formatNumber(row.seen)}
                        <span
                          aria-hidden="true"
                          className="block h-1 w-12 overflow-hidden rounded-full bg-ctp-surface1/60"
                        >
                          <span
                            className="block h-full rounded-full bg-ctp-blue"
                            style={{
                              width: `${results && results.maxSeen > 0 ? (row.seen / results.maxSeen) * 100 : 0}%`,
                            }}
                          />
                        </span>
                      </div>
                    </td>
                    <td className="py-1.5 pl-3 text-right tabular-nums text-ctp-subtext0">
                      {formatNumber(row.titles)}
                    </td>
                    <td className="whitespace-nowrap py-1.5 pl-3 text-right tabular-nums text-ctp-overlay2">
                      {formatFirstSeen(row.firstSeen)}
                    </td>
                    <td className="py-1.5 pl-3 text-right">
                      <button
                        type="button"
                        title="Exclude this word from stats. Restore it from Exclusions."
                        aria-label={`Hide ${row.headword}`}
                        onClick={() => onHideWord(row)}
                        className="text-[11px] text-ctp-overlay2 transition-colors hover:text-ctp-text"
                      >
                        Hide
                      </button>
                    </td>
                  </tr>
                  {expanded ? (
                    <tr className="border-b border-ctp-surface1/50">
                      <td colSpan={6} className="pb-2 pl-5">
                        <div lang="ja" className="flex flex-wrap gap-x-3 gap-y-1">
                          {row.forms.map((form) => (
                            <span key={form.wordId} className="text-ctp-subtext0">
                              {form.word}
                              <span lang="en" className="ml-1 tabular-nums text-ctp-overlay2">
                                {formatNumber(form.seen)}
                              </span>
                            </span>
                          ))}
                          {row.formCount > row.forms.length ? (
                            <span lang="en" className="text-ctp-overlay2">
                              +{formatNumber(row.formCount - row.forms.length)} more
                            </span>
                          ) : null}
                        </div>
                      </td>
                    </tr>
                  ) : null}
                </Fragment>
              );
            })}
          </tbody>
        </table>
        {results && rows.length === 0 ? (
          <div className="py-6 text-center text-xs text-ctp-overlay2">
            {searching
              ? 'No words match that search. Hidden words are managed under Exclusions.'
              : 'No words match these filters.'}
          </div>
        ) : null}
      </div>

      <div className="mt-3 flex items-center justify-between text-[11px] text-ctp-overlay2">
        <span role="status">
          {results
            ? `${formatNumber(firstRow)} to ${formatNumber(lastRow)} of ${formatNumber(total)} ${describeResults(view, searching, hasKnownData)}`
            : 'Loading...'}
        </span>
        <span className="flex items-center gap-3">
          <button
            type="button"
            disabled={page === 0}
            onClick={() => setPage(page - 1)}
            className="transition-colors hover:text-ctp-text disabled:opacity-40"
          >
            Prev
          </button>
          <button
            type="button"
            disabled={lastRow >= total}
            onClick={() => setPage(page + 1)}
            className="transition-colors hover:text-ctp-text disabled:opacity-40"
          >
            Next
          </button>
        </span>
      </div>
    </section>
  );
}
