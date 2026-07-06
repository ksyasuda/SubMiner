type: fixed
area: overlay

- Fixed Kiku manual field grouping freezing the overlay after adding a duplicate card: the field grouping modal now reliably appears above fullscreen mpv on Hyprland/Wayland by re-asserting window placement until the compositor maps the modal window, instead of a single post-show attempt that raced the async map and left the dialog invisible.
- Fixed manual field grouping staying broken after the first attempt: the request resolver is now always cleared once a choice is made or the request is abandoned, so later grouping attempts no longer short-circuit to an instant "Field grouping cancelled".
- Fixed a timed-out or failed field grouping request leaving an orphaned, invisible modal window covering mpv: abandoned requests now tear down the modal window and close the dialog so the overlay recovers immediately.
