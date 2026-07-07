type: fixed
area: overlay

- Fixed subtitle text that Yomitan's parser cannot match (e.g. the truncated volitional in とこ戻ろ…) being rendered as plain, non-interactive text: it was invisible to hover/lookup and excluded from the n+1 word count, which could wrongly mark a sentence as n+1. Unparsed runs are now kept as hoverable tokens matching Yomitan's own segmentation; bracketed SFX/speaker captions and punctuation-only runs are still skipped.
