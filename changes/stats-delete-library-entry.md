type: added
area: stats

- Added a "Delete Entry" action to the stats Library detail view. It removes the whole title in one step — every episode, session, subtitle line, rollup, cover art and vocabulary counts derived from them — and the title disappears from the Library grid. Previously only individual episodes and sessions could be deleted, so a mistakenly created entry had to be cleared episode by episode and still lingered in the library.
- Deletion progress is now shown app-wide instead of only on Overview and Sessions. Every delete (session, session group, episode, and library entry) drives a sweeping bar at the top of the window plus a bottom-right status toast, and both stay visible across tab switches, detail views, and the in-app stats overlay window. The per-tab indicators previously went blank the moment you left the tab that started the delete.
