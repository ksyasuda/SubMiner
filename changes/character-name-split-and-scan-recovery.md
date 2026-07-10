type: fixed
area: overlay

- Character dictionaries now split unspaced AniList native names with MeCab person-name tags and reading validation, while installations without MeCab index both plausible boundaries. Existing snapshots regenerate automatically and upgrade to exact MeCab splits when MeCab becomes available.
- Character-name portraits, highlights, and hover lookup now survive punctuation, unmatched text, context-dependent POS exclusions, and competing generic dictionary matches. Name matches take priority without splitting a strictly longer word such as 空気.
