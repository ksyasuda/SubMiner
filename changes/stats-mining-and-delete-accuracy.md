type: fixed
area: stats

- Fixed sentence mining (single and multi-line) counting each created card twice in session and lifetime stats.
- Fixed deleting a session leaving word and kanji "last seen" dates pointing at the deleted session; affected rows are repaired once in the background on next launch.
- Fixed re-seen words that the vocabulary filter hides still being counted in the new-words charts, so charts and the vocabulary summary now agree. Existing mismatches are repaired by the same background pass.
- Fixed Overview totals, charts, activity calendar, and vocabulary summary staying stale after deleting sessions. Failed refreshes keep the dashboard visible and offer a retry.
