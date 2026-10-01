type: fixed
area: overlay

- Fixed mpv key bindings (input.conf and mpv defaults, e.g. `9`/`0` volume) doing nothing while the overlay had focus when the overlay loaded before SubMiner connected to mpv, common in `mpv.backend: x11` mode. The overlay now reloads mpv's bindings once mpv connects.
