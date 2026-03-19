type: changed
area: overlay

- Excluded interjections and sound-effect tokens from subtitle annotation styling so they no longer inherit misleading lexical highlight treatment while still remaining visible and hoverable as plain subtitle tokens.
- Expanded subtitle annotation noise filtering to also strip annotation metadata from standalone grammar-only helper tokens such as particles, auxiliaries, adnominals, common explanatory endings like `んです` / `のだ`, and merged trailing quote-particle forms like `...って` while keeping them tokenized for hover lookup.
