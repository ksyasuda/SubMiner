export declare function createImmersionDbFixture(dbPath: string): void;
export interface FixtureSessionInput {
    uuid: string;
    videoKey: string;
    animeTitleKey?: string | null;
    animeEpisodesTotal?: number | null;
    startedAtMs: number;
    endedAtMs?: number | null;
    activeWatchedMs?: number;
    cardsMined?: number;
    linesSeen?: number;
    tokensSeen?: number;
    watched?: boolean;
    words?: Array<{
        headword: string;
        word: string;
        reading: string;
        count: number;
    }>;
    applyLifetime?: boolean;
}
/**
 * Insert a session (with video/anime/subtitle-line/occurrence rows) the way
 * the app would have recorded it, optionally crediting lifetime aggregates
 * the way applySessionLifetimeSummary does for locally-recorded sessions.
 */
export declare function insertFixtureSession(dbPath: string, input: FixtureSessionInput): void;
//# sourceMappingURL=immersion-db-fixture.d.ts.map