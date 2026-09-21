type: changed
breaking: true
area: stats

- Reject requests from untrusted browser origins and hosts before stats data, media, or Anki operations run, and require JSON for mutation bodies.
- Load the in-app stats overlay from the local server so it uses the same origin protection as the browser dashboard.
