type: fixed
area: yomitan

- Made subtitle frequency lookup honor Yomitan's configured sort-frequency dictionary so unrelated enabled dictionaries no longer suppress or override the selected frequency source.
- Made subtitle annotation filtering prefer part-of-speech tags from Yomitan lookup results, letting interjections and similar non-card material stay filtered even when MeCab is unavailable.
