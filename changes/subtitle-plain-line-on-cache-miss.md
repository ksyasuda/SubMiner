type: fixed
area: overlay

- Subtitle lines no longer wait for tokenization to finish before appearing, even when the previous line is still being processed. On a tokenization cache miss, the plain line is shown immediately at its cue time and upgrades in place once tokens and annotations are ready; stale results cannot replace newer cues, and the basic plain-text websocket no longer receives a duplicate event for the annotation-only upgrade.
