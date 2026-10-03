type: fixed
area: stats

- Fixed sentence mining (single and multi-line) counting each created card twice in session and lifetime stats.
- Fixed word and kanji dates after session deletion. Historical "first seen" dates and counts survive retention pruning; new occurrences record observation time so deletion can update dates exactly. Legacy earlier dates are preserved when the remaining history cannot prove their origin. Stale "last seen" dates are repaired once in the background on next launch.
- Fixed re-seen words that the vocabulary filter hides still being counted in the new-words charts, so charts and the vocabulary summary now agree. Existing mismatches are repaired by the same background pass.
- Fixed Overview totals, charts, activity calendar, and vocabulary summary staying stale after deleting sessions. Failed refreshes keep the dashboard visible and offer a retry.
