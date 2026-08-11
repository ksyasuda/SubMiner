/*
 * Thresholds that decide when a run of repeated subtitle events is one animation.
 *
 * Three consumers have to agree on these numbers or the same karaoke line is one cue in
 * the sidebar and two hundred in the stats: the file-level cue dedup
 * (`subtitle-cue-dedup`), the live gate that decides what immersion stats record
 * (`subtitle-line-dedup-gate`), and the retroactive database cleanup
 * (`immersion-tracker/duplicate-line-cleanup`).
 */

/**
 * Back-to-back frames of the same animation are authored flush against each other; a
 * tiny tolerance absorbs the centisecond rounding of the ASS timestamp format.
 */
export const DUPLICATE_CUE_GAP_TOLERANCE_SECONDS = 0.05;

/**
 * A burst is a *sequence*. Two adjacent events are two events, not an animation --
 * characters do repeat each other, and a repeated line can legitimately be short.
 */
export const MIN_BURST_EVENTS = 3;

/**
 * Real dialogue holds on screen for about a second, so a run with a couple of much
 * shorter events among them looks like frames. Used only alongside authoring evidence.
 */
export const ANIMATION_FRAME_MAX_SECONDS = 0.3;

/** A karaoke run usually ends on a long "hold" frame, so not every event is short. */
export const MIN_TAGGED_BURST_FRAMES = 2;

/**
 * SRT and VTT carry no authoring metadata at all, so timing is the only signal available
 * -- which makes it the easiest one to get wrong. ASS->SRT conversion leaves frames at
 * ~0.04s, well under any real utterance, and a burst leaves many of them behind. Both
 * bounds are deliberately far stricter than the ASS path: a run of ordinary short lines
 * (`えっ` traded between characters) must not clear them.
 */
export const TIMING_ONLY_FRAME_MAX_SECONDS = 0.1;
export const MIN_TIMING_ONLY_FRAMES = 5;
