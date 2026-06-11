<!-- read_when: changing overlay notification hover, macOS mouse passthrough, or notification actions -->

# macOS Notification Hover Stability Design

Status: approved
Date: 2026-06-09

## Problem

On macOS, hovering a character dictionary build notification can make the card flicker and slide as
if it is hiding, then snap back. The likely trigger is the notification stack changing the overlay
window's mouse-passthrough state for a progress card that has no user action.

## Chosen Approach

Keep non-action overlay notifications visually stable and click-through on hover. Only notifications
with explicit actions should request interactive overlay input. The notification history panel keeps
its existing interactive behavior.

This avoids a macOS mouseenter/mouseleave passthrough loop for passive progress cards while
preserving clickable notification actions.

## Checks

- Add a renderer regression test for passive notification hover.
- Keep action-bearing notification cards interactive.
- Run the targeted overlay notification and mouse-ignore tests.
