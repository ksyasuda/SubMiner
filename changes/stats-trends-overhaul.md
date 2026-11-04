type: changed
area: stats

- Rebuilt the Trends tab: summary totals with period comparison, a Learning section (new words per 1,000 seen, new words by frequency rank), a day and hour heatmap, a title timeline, and a single running-total chart with a metric switcher. These replace the seven cumulative charts, the Sessions chart, the day-of-week and hour charts, the four overlapping per-title charts, and the top titles chart.
- Added week grouping, custom date ranges, a title filter for the whole page, and click-through from charts to sessions and from titles to their library page. Range and grouping are remembered, and changing them no longer blanks the page or loses your scroll position.
- Fixed the Library section including days outside the selected range, so its totals now match Activity.
- Fixed episode counts: Trends now counts finished episodes, each once, instead of counting a video again for every day it was opened.
- New word counts across Trends and Vocabulary now count dictionary words, so a word's conjugations count once. Totals are lower than before.
- Rate charts no longer plot idle days as zero or spike on very short days: they show gaps when there is too little data, and daily views use a rolling 7 days.
- Fixed clipped chart axis labels and month labels that read like dates, and the All range ignoring Show empty days.
- Rebuilt the Vocabulary tab: one word table with Worth mining, All words, Across titles, and Newest views, search across every word, form, and reading, sortable columns, and inline Hide. It replaces three lists that could only see your 500 most seen word forms. Also added known words by frequency band, a vocabulary growth chart, and a kanji view covering every kanji, with a filter for kanji outside your known words.
- Fixed word counts and readings: lists and the word detail panel now total every form of a word and show the dictionary form's reading, so 見る no longer shows みてる. The detail panel loads example lines straight away and lists a word's forms.
- Fixed names slipping past Hide Names and the Trends new-word counts when tagged with more than one part of speech.
- Stats now says when a word or line can't be mined because it came from a stream, in Search and the word panel, instead of showing Unavailable buttons. Lines you can mine are listed first. The word panel counts affected lines, and the word table marks words with only streamed lines as Stream only.
- Fixed stats treating the day as one earlier in time zones ahead of UTC, which affected today's totals, the custom date picker, and first-seen dates.
