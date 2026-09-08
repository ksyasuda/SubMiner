type: changed
area: sync

- Sync uses compressed, incremental rsync transfers on compatible macOS and Linux machines, caching the last received snapshot per peer to reduce traffic on subsequent syncs.
- Machines without compatible rsync, including Windows endpoints, automatically use compressed scp transfers.
