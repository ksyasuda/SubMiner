type: changed
area: startup

- Ordered startup OSD messages so tokenization loads first, annotation loading appears next if still pending, and character dictionary sync progress waits until annotation loading finishes.
