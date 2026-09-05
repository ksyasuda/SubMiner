type: fixed
area: stats

- Stats server startup now reports port conflicts without crashing SubMiner, shares concurrent startup requests, and finishes closing before its tracker is destroyed.
- Stopping during background startup cancels the pending server, and forced application exit attempts to finalize stats even when an HTTP request is still running.
