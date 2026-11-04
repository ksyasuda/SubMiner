type: added
area: jimaku

- Added `jimaku.autoOpenWhenPrimarySubtitleMissing` (off by default). When a local video loads without a Japanese (primary-language) subtitle track, SubMiner pauses playback and opens the Jimaku modal; playback resumes when the modal closes.
- Opening Jimaku while it is already open now reloads the title and episode if the video changed, instead of keeping the previous episode's search.
