type: fixed
area: overlay

- Kanji-bearing nouns that MeCab tags as non-independent (非自立) — e.g. 日 in いい日だったな, 点, 以外 — now keep frequency/JLPT highlighting and count toward vocabulary stats. Yomitan segments them as standalone vocabulary tokens, so the MeCab POS filter only suppresses kana grammar nouns (こと, もの, とき) it was meant for.
