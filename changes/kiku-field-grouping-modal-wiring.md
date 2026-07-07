type: fixed
area: overlay

- Fixed the Kiku field grouping modal misbehaving on Hyprland/Wayland while every other modal worked: its main-process wiring silently dropped the modal-open acknowledgement/retry, failure teardown, and warn-logging callbacks (they are optional, so the incomplete forward compiled cleanly), and it skipped the overlay prerequisites every other modal runs before opening. The field grouping modal now opens through the same path as the other modals — preparing the overlay runtime and visible overlay window first, then acknowledging/retrying the dedicated modal window — so it appears above and works over fullscreen mpv like the rest.
