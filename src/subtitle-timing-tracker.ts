/*
 * SubMiner - Subtitle mining overlay for mpv
 * Copyright (C) 2024 sudacode
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

interface TimingEntry {
  startTime: number;
  endTime: number;
  timestamp: number;
}

interface HistoryEntry {
  displayText: string;
  timingKey: string;
  startTime: number;
  endTime: number;
  timestamp: number;
}

export class SubtitleTimingTracker {
  private timings = new Map<string, TimingEntry>();
  private history: HistoryEntry[] = [];
  private readonly maxHistory = 200;
  private readonly ttlMs = 5 * 60 * 1000;
  private cleanupInterval: ReturnType<typeof setInterval> | null = null;

  constructor() {
    this.startCleanup();
  }

  recordSubtitle(text: string, startTime: number, endTime: number): void {
    const normalizedText = this.normalizeText(text);
    if (!normalizedText) return;

    const displayText = this.prepareDisplayText(text);
    const timingKey = normalizedText;

    this.timings.set(timingKey, {
      startTime,
      endTime,
      timestamp: Date.now(),
    });

    // Check for duplicate of most recent entry (deduplicate adjacent repeats)
    const lastEntry = this.history[this.history.length - 1];
    if (lastEntry && lastEntry.timingKey === timingKey) {
      // Update timing to most recent occurrence
      lastEntry.startTime = startTime;
      lastEntry.endTime = endTime;
      lastEntry.timestamp = Date.now();
      return;
    }

    this.history.push({
      displayText,
      timingKey,
      startTime,
      endTime,
      timestamp: Date.now(),
    });

    // Prune history if too large
    if (this.history.length > this.maxHistory) {
      this.history = this.history.slice(-this.maxHistory);
    }
  }

  findTiming(text: string): { startTime: number; endTime: number } | null {
    const normalizedText = this.normalizeText(text);
    if (!normalizedText) return null;

    const entry = this.timings.get(normalizedText);
    if (!entry) {
      return this.findFuzzyMatch(normalizedText);
    }

    return {
      startTime: entry.startTime,
      endTime: entry.endTime,
    };
  }

  /**
   * Get recent subtitle blocks in chronological order.
   * Returns the last `count` subtitle events (oldest → newest).
   * Blocks preserve internal line breaks and are joined with blank lines.
   */
  getRecentBlocks(count: number): string[] {
    if (count <= 0) return [];
    if (count > this.history.length) {
      count = this.history.length;
    }
    return this.history.slice(-count).map((entry) => entry.displayText);
  }

  /**
   * Get display text for the most recent subtitle.
   */
  getCurrentSubtitle(): string | null {
    const lastEntry = this.history[this.history.length - 1];
    return lastEntry ? lastEntry.displayText : null;
  }

  private findFuzzyMatch(
    text: string,
  ): { startTime: number; endTime: number } | null {
    let bestMatch: TimingEntry | null = null;
    let bestScore = 0;

    for (const [key, entry] of this.timings.entries()) {
      const score = this.calculateSimilarity(text, key);
      if (score > bestScore && score > 0.7) {
        bestScore = score;
        bestMatch = entry;
      }
    }

    if (bestMatch) {
      return {
        startTime: bestMatch.startTime,
        endTime: bestMatch.endTime,
      };
    }

    return null;
  }

  private calculateSimilarity(a: string, b: string): number {
    const longer = a.length > b.length ? a : b;
    const shorter = a.length > b.length ? b : a;

    if (longer.length === 0) return 1;

    const editDistance = this.getEditDistance(longer, shorter);
    return (longer.length - editDistance) / longer.length;
  }

  private getEditDistance(longer: string, shorter: string): number {
    const costs: number[] = [];
    for (let i = 0; i <= shorter.length; i++) {
      let lastValue = i;
      for (let j = 1; j <= longer.length; j++) {
        let newValue = costs[j - 1] || 0;
        if (longer.charAt(j - 1) !== shorter.charAt(i - 1)) {
          newValue = Math.min(Math.min(newValue, lastValue), costs[j] || 0) + 1;
        }
        costs[j - 1] = lastValue;
        lastValue = newValue;
      }
      costs[shorter.length] = lastValue;
    }
    return costs[shorter.length] || 0;
  }

  private normalizeText(text: string): string {
    return text
      .replace(/\\N/g, " ")
      .replace(/\\n/g, " ")
      .replace(/\n/g, " ")
      .replace(/{[^}]*}/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  private prepareDisplayText(text: string): string {
    // Convert ASS/SSA newlines to real newlines, strip tags
    return text
      .replace(/\\N/g, "\n")
      .replace(/\\n/g, "\n")
      .replace(/{[^}]*}/g, "")
      .trim();
  }

  private startCleanup(): void {
    this.cleanupInterval = setInterval(() => {
      this.cleanup();
    }, 60000);
  }

  cleanup(): void {
    const now = Date.now();
    // Clean up old timing entries
    for (const [key, entry] of this.timings.entries()) {
      if (now - entry.timestamp > this.ttlMs) {
        this.timings.delete(key);
      }
    }
    // Clean up old history entries
    this.history = this.history.filter(
      (entry) => now - entry.timestamp <= this.ttlMs,
    );
  }

  destroy(): void {
    if (this.cleanupInterval) {
      clearInterval(this.cleanupInterval);
      this.cleanupInterval = null;
    }
    this.timings.clear();
    this.history = [];
  }
}
