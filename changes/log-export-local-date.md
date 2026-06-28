type: fixed
area: logs

- Fixed log filenames to use the user's local date so log export includes the current logs instead of stale prior-day logs around UTC midnight.
- Expanded log export redaction to mask common PII and secrets, including IP addresses, emails, auth/cookie headers, yt-dlp cookie arguments, URL credentials, compound token/key/password fields, and signed YouTube media URL query strings.
