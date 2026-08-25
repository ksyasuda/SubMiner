type: fixed
area: stats

- Immersion statistics storage now applies its SQLite busy timeout before WAL setup, avoiding transient database-lock failures when worker connections overlap.
