type: changed
area: sync

- Sync uses compressed, incremental rsync transfers on compatible macOS and Linux machines, caching the last received snapshot per peer to reduce traffic on subsequent syncs. Cache helpers work through the launcher; older apps and launchers fall back to compressed transfers without an upload cache.
- Machines without compatible rsync, including Windows endpoints, automatically use compressed scp transfers.
- Rsync explicitly uses SSH and aborts transfers that exceed 30 minutes before merging.
