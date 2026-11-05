import { readBooleanPreference, readStringPreference } from '../../lib/preference-storage';
import type {
  StatsDictionaryWordPos,
  StatsDictionaryWordQuery,
  StatsDictionaryWordSort,
} from '../../types/stats';

export const WORD_VIEWS = ['worth', 'all', 'titles', 'newest'] as const;
export type WordView = (typeof WORD_VIEWS)[number];

export const WORD_PAGE_SIZE = 25;
// Beyond this rank a word is rarely worth prioritizing, and sorting by rank alone
// surfaces tokenizer artifacts, so Worth mining ranks by how often you meet a word.
const WORTH_MINING_MAX_RANK = 10_000;
const ACROSS_TITLES_MIN_TITLES = 2;

export const HIDE_KNOWN_KEY = 'subminer-stats-vocabulary-hide-known';
export const HIDE_KANA_KEY = 'subminer-stats-vocabulary-hide-kana';
export const HIDE_NAMES_KEY = 'subminer-stats-vocabulary-hide-names';
export const WORD_VIEW_KEY = 'subminer-stats-vocabulary-view';

export function loadWordView(): WordView {
  const raw = readStringPreference(WORD_VIEW_KEY);
  return WORD_VIEWS.find((view) => view === raw) ?? 'worth';
}

export function loadHideNames(): boolean {
  return readBooleanPreference(HIDE_NAMES_KEY, true);
}

export interface WordSort {
  sort: StatsDictionaryWordSort;
  descending: boolean;
}

const VIEW_DEFAULT_SORT: Record<WordView, WordSort> = {
  worth: { sort: 'seen', descending: true },
  all: { sort: 'seen', descending: true },
  titles: { sort: 'titles', descending: true },
  newest: { sort: 'firstSeen', descending: true },
};

export function defaultSortForView(view: WordView): WordSort {
  return VIEW_DEFAULT_SORT[view];
}

/** Clicking a sorted column flips it; a new column starts on its most useful end. */
export function nextSort(current: WordSort, column: StatsDictionaryWordSort): WordSort {
  if (current.sort === column) return { sort: column, descending: !current.descending };
  return { sort: column, descending: column !== 'rank' && column !== 'headword' };
}

export interface WordExplorerState {
  view: WordView;
  search: string;
  hideKnown: boolean;
  hideNames: boolean;
  hideKana: boolean;
  partOfSpeech: StatsDictionaryWordPos | null;
  sort: WordSort;
  page: number;
}

/**
 * Turns the controls into a server query. Searching looks through everything, so
 * typing a word you know or a name still finds it whatever the view and filters say.
 */
export function buildWordQuery(state: WordExplorerState): StatsDictionaryWordQuery {
  const paging = {
    ...state.sort,
    limit: WORD_PAGE_SIZE,
    offset: state.page * WORD_PAGE_SIZE,
  };
  const search = state.search.trim();
  if (search) {
    return { ...paging, search, hideKnown: false, hideNames: false, hideKana: false };
  }
  const isWorth = state.view === 'worth';
  return {
    ...paging,
    hideKnown: isWorth || state.hideKnown,
    hideNames: state.hideNames,
    hideKana: state.hideKana,
    partOfSpeech: state.partOfSpeech ?? undefined,
    maxRank: isWorth ? WORTH_MINING_MAX_RANK : undefined,
    minTitles: state.view === 'titles' ? ACROSS_TITLES_MIN_TITLES : undefined,
  };
}
