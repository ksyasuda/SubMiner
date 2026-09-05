type: internal
area: runtime

- Move Linux overlay mode state and window replacement into one runtime owner and cancel pending replacements during teardown.
- Remove the protocol-handler forwarding builder and redundant lifecycle callback wrappers.
