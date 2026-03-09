type: fixed
area: dictionary

- Keep character dictionary auto-sync non-blocking during startup by letting snapshot/build work run in parallel and delaying only the Yomitan import/settings phase until current-media tokenization is already ready.
