type: added
area: overlay

- Added a controller layout for the media timing review dialog covering every button, with L1 / R1 + D-pad combos for the clip start and end, and D-pad actions that repeat while held. Each button shows its controller shortcut while controller support is on, and the first press in a newly opened dialog already counts. The controller's usual actions pause while the dialog is open, including in the hidden subtitle overlay behind it.
- Added `controller.faceButtonLabels` (`xbox` or `nintendo`) so fixed controller layouts and the `Alt+C` button names follow the letters printed on the pad. Use `nintendo` for controllers with A on the right and B at the bottom.
