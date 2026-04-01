type: fixed
area: overlay

- Fixed Kiku duplicate grouping to reuse duplicate note IDs from both generic sentence-card creation and Yomitan popup mining instead of running extra duplicate scans after add.
- Fixed the Yomitan popup mining flow to add cards in the background while keeping the stock popup progress feedback, then pause playback and close the lookup popup before the Kiku merge modal opens.
- Fixed configured subtitle-jump keybindings so backward and forward subtitle seeks keep playback paused when invoked from a paused state.
