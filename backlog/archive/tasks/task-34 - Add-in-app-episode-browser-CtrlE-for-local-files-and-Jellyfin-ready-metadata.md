---
id: TASK-34
title: >-
  Add in-app episode browser (Ctrl+E) for local files and Jellyfin-ready
  metadata
status: To Do
assignee: []
created_date: '2026-02-13 22:12'
updated_date: '2026-02-13 22:13'
labels: []
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Implement an in-app episode browser invoked by Ctrl+E when mpv is open and connected to the Electron UI. The viewer should use the currently supplied directory (if app was launched with one) or fallback to the parent directory of the currently playing video. It should enumerate all available video files in target directory and sort them deterministically, then display them in a polished list/gallery with thumbnails in the Electron UI. Thumbnail behavior should prioritize existing matching images in the directory; otherwise generate thumbnails asynchronously in the background and update the modal as they become available. The same menu infrastructure should be shared with the planned Jellyfin integration and designed so it can display additional Jellyfin-sourced metadata without UI rewrites. It should also support launching/using an alternate picker mode compatible with external launcher UX patterns (e.g., via fzf/rofi), showing the same episode list/metadata in those contexts.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Pressing Ctrl+E in connected mode opens an episode viewer modal and closes/overlays correctly in the Electron app.
- [ ] #2 Episode browser source directory is resolved from CLI-provided path when present, else from currently playing video's parent directory.
- [ ] #3 Browser enumerates all supported video files in target directory and sorts them deterministically (e.g., natural/season-episode aware when possible).
- [ ] #4 Episode list/gallery renders a consistent, usable layout with title, position/index, and thumbnail placeholders.
- [ ] #5 If a thumbnail image file with matching name exists in the directory, it is used without background generation.
- [ ] #6 For files without matching thumbnails, background thumbnail generation runs and updates entries as images become available in the modal.
- [ ] #7 Thumbnail generation is cancellable/abortable and does not block UI interaction.
- [ ] #8 The same episode viewer component/path can be reused by Jellyfin integration and can accept extended metadata payloads (e.g., show title, season/episode, runtime, description).
- [ ] #9 `fzf`/`rofi`-compatible episode picker mode is available so the same episode set can be browsed outside the Electron modal using either backend output format.
- [ ] #10 Episode item ordering, labels, and metadata shown in fzf/rofi mode match the in-app sorting and identity scheme.
- [ ] #11 Both picker modes (Electron modal and fzf/rofi) resolve source directory using the same CLI/parent-of-current-video precedence rules.
- [ ] #12 When invoked through fzf/rofi, the selected episode can be played and returns focus/flow safely to the Electron-mpv workflow.
<!-- AC:END -->

## Definition of Done

<!-- DOD:BEGIN -->

- [ ] #1 Feature supports Ctrl+E invocation from connected mpv session.
- [ ] #2 Directory fallback behavior is implemented and validated with both passed-in and default paths.
- [ ] #3 Video listing excludes unsupported formats and is correctly sorted.
- [ ] #4 Existing local thumbnail matching works for common image/video naming patterns.
- [ ] #5 Background thumbnail generation works asynchronously and updates UI in-place.
- [ ] #6 UI is ready for Jellyfin metadata fields and does not require structural rewrite for server-provided data.
- [ ] #7 No regressions in existing mpv-Electron controls for standard playback.
<!-- DOD:END -->
