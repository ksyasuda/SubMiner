type: fixed
area: overlay

- Cancel pending Linux overlay window replacements during teardown so a delayed close callback cannot reopen the overlay.
