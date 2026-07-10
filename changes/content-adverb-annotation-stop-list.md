type: fixed
area: overlay

- Frequency/JLPT highlighting and vocabulary stats now include content adverbs such as 確かに and やはり, plus kanji-bearing nouns that MeCab tags as non-independent such as 日, 点, and 以外. The noise filters still suppress interjections, pronouns, grammar fragments, and kana grammar nouns.
- Lexicalized kana expressions such as `かといって` retain their frequency annotations when their merged MeCab parts include particles.
