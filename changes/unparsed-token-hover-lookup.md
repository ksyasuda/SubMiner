type: fixed
area: overlay

- Fixed subtitle text that Yomitan's parser cannot match (truncated inflections like the volitional in とこ戻ろ…, elongation runs like ぅ～/ぉ〜) being rendered as plain, non-interactive text with no hover lookup. Unparsed runs are now kept as hoverable tokens matching Yomitan's own segmentation, while being fully ignored by frequency/JLPT highlighting, the N+1 candidate math, and vocabulary stats. Bracketed SFX/speaker captions and punctuation-only runs are still skipped.
