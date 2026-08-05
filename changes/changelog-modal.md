type: added
area: overlay

- Added an in-app changelog modal, opened from the tray ("View Changelog") or the "What's New" button on the update-available notification, which now stays on screen so "Update" is still reachable after reading the notes. It renders inside the player bounds when a video is playing and in its own window otherwise, the same as the help modal.
- The changelog is fetched from the newest published release, so release notes for versions newer than the installed build are visible; a failed download falls back to the changelog bundled with the install and says so in the modal.
- Versions are foldable: the current `0.x` line is expanded and older lines are folded, matching the docs-site changelog. A badge marks the installed version and newer versions are tagged "New".
- Keyboard: `J`/`K` or arrows move between versions, `Enter` folds/unfolds, `R` refetches, `Esc` closes.
