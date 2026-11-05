type: fixed
area: overlay

- The controller settings window (`Alt+C`) no longer edits a controller it cannot see. A newly opened window cannot see any controller until you press one of its buttons (browsers hide a pad from a new window until then); it used to show and edit the global defaults meanwhile, and its status could stay stuck on "No controllers detected". Now it shows the bindings of the controller the overlay is using (or the saved one), labelled as a preview, and Learn, Clear, Reset and the button style are locked until the press.
- With several controllers connected, `Alt+C` now selects the one you just pressed, and Learn listens to that controller.
