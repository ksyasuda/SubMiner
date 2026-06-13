type: internal
area: runtime

- Split main-process runtime wiring into focused modules without changing user-facing behavior.
- Hardened split runtime helpers against stale background stats daemon PIDs, stalled subtitle extraction, and dropped async errors.
