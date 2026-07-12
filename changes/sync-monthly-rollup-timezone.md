type: fixed
area: sync

- Fixed stats sync copying a remote daily rollup that should have been recomputed when the day was the 1st of a month and the machine was on a negative UTC offset (the Americas), which could leave that day's totals wrong after a sync. The rollup day is now read back at local noon instead of UTC midnight, so it always resolves to the correct civil month.
