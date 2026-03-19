type: fixed
area: stats

- `subminer stats -b` now runs as a standalone background stats daemon instead of reusing the main SubMiner app process, so the overlay app can still be launched separately for normal video watching.
- Dashboard word mining still works against the background daemon by using a short-lived hidden helper for the Yomitan add-note flow.
