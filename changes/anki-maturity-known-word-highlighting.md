type: added
area: overlay

- Known-word subtitle highlights can now be colored by Anki card maturity (new, learning, young, mature) like asbplayer. Enable with `ankiConnect.knownWords.maturityEnabled`; the mature interval threshold (`matureThresholdDays`, default 21) and the four tier colors (`subtitleStyle.knownWordMaturityColors`) are configurable, and a runtime option toggles it in-session. The session help color legend shows the four tier colors while maturity highlighting is on.
- Tiers follow Anki's own card counts: the interval tiers exclude cards in the learning/relearning queue, so a lapsed card shows the learning color instead of young (its interval is reset to at least 1 day, which previously made the learning tier unreachable). A note with a mature card alongside a relearning card still shows mature.
