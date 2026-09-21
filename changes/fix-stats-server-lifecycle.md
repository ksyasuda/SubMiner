type: fixed
area: stats

- Stats server startup reports port conflicts without crashing SubMiner, shares concurrent startup requests, and shows in-app startup errors through configured status notifications.
- Background stop cancels pending background startup without disconnecting foreground-only dashboards. Shutdown bounds the wait for active HTTP requests and awaits tracker finalization before exit, with a deadline for forced application exit.
