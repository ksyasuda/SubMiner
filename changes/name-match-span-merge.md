type: fixed
area: overlay

- Fixed character-name annotations dropping for an entire subtitle line when it contained any chunk the dictionary scanner could not match (e.g. an interjection like やほっ before a name): scanner metadata is now merged per token into the parseText segmentation instead of being discarded on any span mismatch.
