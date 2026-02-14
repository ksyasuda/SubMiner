---
id: TASK-43
title: Add community subtitle timing database for shared sync corrections
status: To Do
assignee: []
created_date: '2026-02-14 02:13'
labels:
  - feature
  - community
  - subtitles
  - sync
dependencies: []
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Allow users to share their subtitle timing corrections to a community database, so other users watching the same video file get pre-synced subtitles automatically.

## Motivation
Subtitle synchronization (alass/ffsubsync) is one of the most friction-heavy steps in the mining workflow. Users spend time syncing subtitles that someone else has already synced for the exact same video. A shared database of timing corrections keyed by video file hash would eliminate redundant work.

## Design
1. **Video identification**: Use a partial file hash (first + last N bytes, or a media fingerprint) to identify video files without uploading content
2. **Timing data**: Store the timing offset/warp parameters produced by alass/ffsubsync, not the full subtitle file
3. **Upload flow**: After a successful sync, offer to share the timing correction (opt-in)
4. **Download flow**: Before syncing, check the community database for existing corrections for the current video hash
5. **Trust model**: Simple upvote/downvote on corrections; show number of users who confirmed a correction works

## Technical considerations
- Backend could be a simple REST API with a lightweight database (or even a GitHub-hosted JSON/SQLite file for v1)
- Privacy: only file hashes and timing parameters are shared, never video content or personal data
- Subtitle source (jimaku entry ID) can serve as an additional matching key
- Rate limiting and abuse prevention needed for public API
- Could integrate with existing jimaku modal flow

## Phasing
- v1: Local export/import of timing corrections (share as files)
- v2: Optional cloud sync with community database
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Video files are identified by content hash without uploading video data.
- [ ] #2 Timing corrections (offset/warp parameters) can be exported and shared.
- [ ] #3 Before syncing, the app checks for existing community corrections for the current video.
- [ ] #4 Upload of timing data is opt-in with clear privacy disclosure.
- [ ] #5 Downloaded corrections are applied automatically or with one-click confirmation.
- [ ] #6 Trust signal (confirmation count) is shown for community corrections.
<!-- AC:END -->
