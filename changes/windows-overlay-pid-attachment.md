type: fixed
area: overlay

- Windows overlay attachment now identifies mpv through its IPC process ID, removing the dependency on PowerShell command-line lookups and allowing the IPC socket to be configured in `mpv.conf`.
- Revalidate cached Windows mpv process IDs and expire them after sustained IPC failures so PID reuse cannot keep the overlay attached to an unrelated window.
- Diagnose stalled Windows overlays at the default warn log level, including tracking/PID state, renderer readiness, opacity, native binding errors, and recovery, with repeated warnings rate limited.
