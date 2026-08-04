type: fixed
area: overlay

- Subtitle lines no longer wait for tokenization to finish before appearing. On a tokenization cache miss, the plain line is shown immediately at its cue time and upgrades in place once tokens and annotations are ready; cached lines still appear fully annotated on time as before.
