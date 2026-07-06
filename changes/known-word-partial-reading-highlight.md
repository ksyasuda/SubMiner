type: fixed
area: overlay

- Fixed never-mined compound words (e.g. 待ち合わせてる) being highlighted green as known: subtitle tokens now carry complete readings instead of kanji-only furigana joins, and the known-word reading fallback rejects readings that don't cover the token surface. Stored word readings in the stats database are no longer truncated for new lines.
