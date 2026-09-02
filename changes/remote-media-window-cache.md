type: changed
area: mining

- Mining a card from a remote stream (Jellyfin and other HTTP sources) now downloads the clip window once into a temporary file and reuses it for the timing review waveform, audio preview, audio extraction, and screenshot, instead of re-fetching the stream for every step. The temporary window is removed after ten minutes without use or on exit.
