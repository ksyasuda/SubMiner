type: added
area: stats

- Library: duplicate cards for the same show can be combined. Press "Select" above the library grid, tick the cards, and use "Merge Selected"; the dialog picks which entry to keep and moves every episode onto it. Sessions, mined cards, and watch time are preserved, emptied entries disappear, and remembered title aliases keep future episodes on the merged card.
- Library: episodes can be reassigned to another entry from the "→" button on an episode row, which is the fix when one file lands under a stray title (for example an episode name parsed as the series). Manual assignments survive later filename parsing, Jellyfin refreshes, and season repair, and emptying an entry this way removes it and returns to the grid.
- Library: exact AniList title matches with compatible seasons fold duplicate cards automatically. Fuzzy same-AniList matches appear as dismissible "Possible duplicate" reviews instead of changing the library without confirmation.
