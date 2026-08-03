type: changed
area: subsync

- The subsync modal now lets you pick both sides of an alass run: the reference subtitle (correct timing, defaults to the loaded secondary subtitle track) and the out-of-sync subtitle that gets retimed (defaults to the active primary track).
- alass can now use the loaded video file itself as the reference (audio-based, local files only). It is offered in the reference list but is never the default.
- The out-of-sync subtitle picker also applies to ffsubsync, so a track other than the active primary one can be retimed.
- Retiming the secondary subtitle track now reloads the synced result back into the secondary slot and leaves the primary track selected, instead of replacing the primary subtitle.
