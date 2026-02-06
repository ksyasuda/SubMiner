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

import { AnkiConnectClient } from "./anki-connect";
import { SubtitleTimingTracker } from "./subtitle-timing-tracker";
import { MediaGenerator } from "./media-generator";
import {
  AnkiConnectConfig,
  KikuDuplicateCardInfo,
  KikuFieldGroupingChoice,
  MpvClient,
  NotificationOptions,
} from "./types";

interface NoteInfo {
  noteId: number;
  fields: Record<string, { value: string }>;
}

export class AnkiIntegration {
  private client: AnkiConnectClient;
  private mediaGenerator: MediaGenerator;
  private timingTracker: SubtitleTimingTracker;
  private config: AnkiConnectConfig;
  private pollingInterval: ReturnType<typeof setInterval> | null = null;
  private previousNoteIds = new Set<number>();
  private initialized = false;
  private backoffMs = 200;
  private maxBackoffMs = 5000;
  private nextPollTime = 0;
  private mpvClient: MpvClient;
  private osdCallback: ((text: string) => void) | null = null;
  private notificationCallback:
    | ((title: string, options: NotificationOptions) => void)
    | null = null;
  private updateInProgress = false;
  private fieldGroupingCallback:
    | ((data: {
        original: KikuDuplicateCardInfo;
        duplicate: KikuDuplicateCardInfo;
      }) => Promise<KikuFieldGroupingChoice>)
    | null = null;

  constructor(
    config: AnkiConnectConfig,
    timingTracker: SubtitleTimingTracker,
    mpvClient: MpvClient,
    osdCallback?: (text: string) => void,
    notificationCallback?: (
      title: string,
      options: NotificationOptions,
    ) => void,
    fieldGroupingCallback?: (data: {
      original: KikuDuplicateCardInfo;
      duplicate: KikuDuplicateCardInfo;
    }) => Promise<KikuFieldGroupingChoice>,
  ) {
    this.config = {
      url: "http://127.0.0.1:8765",
      pollingRate: 3000,
      audioField: "ExpressionAudio",
      imageField: "Picture",
      sentenceField: "Sentence",
      generateAudio: true,
      generateImage: true,
      imageType: "static",
      imageFormat: "jpg",
      overwriteAudio: true,
      overwriteImage: true,
      mediaInsertMode: "append",
      audioPadding: 0.5,
      fallbackDuration: 3.0,
      miscInfoPattern: "[SubMiner] %f (%t)",
      notificationType: "osd",
      imageQuality: 92,
      animatedFps: 10,
      animatedMaxWidth: 640,
      animatedCrf: 35,
      autoUpdateNewCards: true,
      maxMediaDuration: 30,
      audioCardField: "IsAudioCard",
      ...config,
    };

    this.client = new AnkiConnectClient(this.config.url!);
    this.mediaGenerator = new MediaGenerator();
    this.timingTracker = timingTracker;
    this.mpvClient = mpvClient;
    this.osdCallback = osdCallback || null;
    this.notificationCallback = notificationCallback || null;
    this.fieldGroupingCallback = fieldGroupingCallback || null;
  }

  start(): void {
    if (this.pollingInterval) {
      this.stop();
    }

    console.log(
      "Starting AnkiConnect integration with polling rate:",
      this.config.pollingRate,
    );
    this.poll();
  }

  stop(): void {
    if (this.pollingInterval) {
      clearInterval(this.pollingInterval);
      this.pollingInterval = null;
    }
    console.log("Stopped AnkiConnect integration");
  }

  private poll(): void {
    this.pollOnce();
    this.pollingInterval = setInterval(() => {
      this.pollOnce();
    }, this.config.pollingRate);
  }

  private async pollOnce(): Promise<void> {
    if (this.updateInProgress) return;
    if (Date.now() < this.nextPollTime) return;

    this.updateInProgress = true;
    try {
      const query = this.config.deck
        ? `"deck:${this.config.deck}" added:1`
        : "added:1";
      const noteIds = (await this.client.findNotes(query, {
        maxRetries: 0,
      })) as number[];
      const currentNoteIds = new Set(noteIds);

      if (!this.initialized) {
        this.previousNoteIds = currentNoteIds;
        this.initialized = true;
        console.log(
          `AnkiConnect initialized with ${currentNoteIds.size} existing cards`,
        );
        this.backoffMs = 200;
        return;
      }

      const newNoteIds = Array.from(currentNoteIds).filter(
        (id) => !this.previousNoteIds.has(id),
      );

      if (newNoteIds.length > 0) {
        console.log("Found new cards:", newNoteIds);

        for (const noteId of newNoteIds) {
          this.previousNoteIds.add(noteId);
        }

        if (this.config.autoUpdateNewCards !== false) {
          for (const noteId of newNoteIds) {
            await this.processNewCard(noteId);
          }
        } else {
          console.log(
            "New card detected (auto-update disabled). Press Ctrl+V to update from clipboard.",
          );
        }
      }

      if (this.backoffMs > 200) {
        console.log("AnkiConnect connection restored");
      }
      this.backoffMs = 200;
    } catch (error) {
      const wasBackingOff = this.backoffMs > 200;
      this.backoffMs = Math.min(this.backoffMs * 2, this.maxBackoffMs);
      this.nextPollTime = Date.now() + this.backoffMs;
      if (!wasBackingOff) {
        console.warn("AnkiConnect polling failed, backing off...");
        this.showStatusNotification("AnkiConnect: unable to connect");
      }
    } finally {
      this.updateInProgress = false;
    }
  }

  private async processNewCard(noteId: number): Promise<void> {
    try {
      const notesInfoResult = await this.client.notesInfo([noteId]);
      const notesInfo = notesInfoResult as unknown as NoteInfo[];
      if (!notesInfo || notesInfo.length === 0) {
        console.warn("Card not found:", noteId);
        return;
      }

      const noteInfo = notesInfo[0];
      const fields = this.extractFields(noteInfo.fields);

      const expressionText = fields.expression || fields.word || "";
      if (!expressionText) {
        console.warn("No expression/word field found in card:", noteId);
        return;
      }

      if (
        this.config.isKiku &&
        this.config.kikuFieldGrouping &&
        this.config.kikuFieldGrouping !== "disabled"
      ) {
        const duplicateNoteId = await this.findDuplicateNote(
          expressionText,
          noteId,
          noteInfo,
        );
        if (duplicateNoteId !== null) {
          if (this.config.kikuFieldGrouping === "auto") {
            await this.handleFieldGroupingAuto(
              duplicateNoteId,
              noteId,
              noteInfo,
              expressionText,
            );
            return;
          } else if (this.config.kikuFieldGrouping === "manual") {
            const handled = await this.handleFieldGroupingManual(
              duplicateNoteId,
              noteId,
              noteInfo,
              expressionText,
            );
            if (handled) return;
          }
        }
      }

      const updatedFields: Record<string, string> = {};
      let updatePerformed = false;

      if (this.config.sentenceField && this.mpvClient.currentSubText) {
        const processedSentence = this.processSentence(
          this.mpvClient.currentSubText,
          fields,
        );
        updatedFields[this.config.sentenceField] = processedSentence;
        updatePerformed = true;
      }

      if (this.config.generateAudio && this.mpvClient) {
        try {
          const audioFilename = this.generateAudioFilename();
          const audioBuffer = await this.generateAudio();

          if (audioBuffer) {
            await this.client.storeMediaFile(audioFilename, audioBuffer);
            const existingAudio =
              noteInfo.fields[this.config.audioField!]?.value || "";
            updatedFields[this.config.audioField!] = this.mergeFieldValue(
              existingAudio,
              `[sound:${audioFilename}]`,
              this.config.overwriteAudio !== false,
            );

            if (this.config.miscInfoField) {
              const miscInfo = this.formatMiscInfoPattern(audioFilename);
              if (miscInfo) {
                updatedFields[this.config.miscInfoField] = miscInfo;
              }
            }

            updatePerformed = true;
          }
        } catch (error) {
          console.error("Failed to generate audio:", (error as Error).message);
          this.showOsdNotification(
            `Audio generation failed: ${(error as Error).message}`,
          );
        }
      }

      let imageBuffer: Buffer | null = null;
      if (this.config.generateImage && this.mpvClient) {
        try {
          const imageFilename = this.generateImageFilename();
          imageBuffer = await this.generateImage();

          if (imageBuffer) {
            await this.client.storeMediaFile(imageFilename, imageBuffer);
            const existingImage =
              noteInfo.fields[this.config.imageField!]?.value || "";
            updatedFields[this.config.imageField!] = this.mergeFieldValue(
              existingImage,
              `<img src="${imageFilename}">`,
              this.config.overwriteImage !== false,
            );

            if (
              this.config.miscInfoField &&
              !updatedFields[this.config.miscInfoField]
            ) {
              const miscInfo = this.formatMiscInfoPattern(imageFilename);
              if (miscInfo) {
                updatedFields[this.config.miscInfoField] = miscInfo;
              }
            }

            updatePerformed = true;
          }
        } catch (error) {
          console.error("Failed to generate image:", (error as Error).message);
          this.showOsdNotification(
            `Image generation failed: ${(error as Error).message}`,
          );
        }
      }

      if (updatePerformed) {
        await this.client.updateNoteFields(noteId, updatedFields);
        console.log("Updated card fields for:", expressionText);
        await this.showNotification(noteId, expressionText);
      }
    } catch (error) {
      if ((error as Error).message.includes("note was not found")) {
        console.warn("Card was deleted before update:", noteId);
      } else {
        console.error("Error processing new card:", (error as Error).message);
      }
    }
  }

  private extractFields(
    fields: Record<string, { value: string }>,
  ): Record<string, string> {
    const result: Record<string, string> = {};
    for (const [key, value] of Object.entries(fields)) {
      result[key.toLowerCase()] = value.value || "";
    }
    return result;
  }

  private processSentence(
    mpvSentence: string,
    noteFields: Record<string, string>,
  ): string {
    if (this.config.highlightWord === false) {
      return mpvSentence;
    }

    const sentenceFieldName =
      this.config.sentenceField?.toLowerCase() || "sentence";
    const existingSentence = noteFields[sentenceFieldName] || "";

    const highlightMatch = existingSentence.match(/<b>(.*?)<\/b>/);
    if (!highlightMatch || !highlightMatch[1]) {
      return mpvSentence;
    }

    const highlightedText = highlightMatch[1];
    const index = mpvSentence.indexOf(highlightedText);

    if (index === -1) {
      return mpvSentence;
    }

    const prefix = mpvSentence.substring(0, index);
    const suffix = mpvSentence.substring(index + highlightedText.length);
    return `${prefix}<b>${highlightedText}</b>${suffix}`;
  }

  private async generateAudio(): Promise<Buffer | null> {
    const mpvClient = this.mpvClient;
    if (!mpvClient || !mpvClient.currentVideoPath) {
      return null;
    }

    const videoPath = mpvClient.currentVideoPath;
    let startTime = mpvClient.currentSubStart;
    let endTime = mpvClient.currentSubEnd;

    if (startTime === undefined || endTime === undefined) {
      const currentTime = mpvClient.currentTimePos || 0;
      const fallback = this.config.fallbackDuration! / 2;
      startTime = currentTime - fallback;
      endTime = currentTime + fallback;
    }

    return this.mediaGenerator.generateAudio(
      videoPath,
      startTime,
      endTime,
      this.config.audioPadding,
    );
  }

  private async generateImage(): Promise<Buffer | null> {
    if (!this.mpvClient || !this.mpvClient.currentVideoPath) {
      return null;
    }

    const videoPath = this.mpvClient.currentVideoPath;
    const timestamp = this.mpvClient.currentTimePos || 0;

    if (this.config.imageType === "avif") {
      let startTime = this.mpvClient.currentSubStart;
      let endTime = this.mpvClient.currentSubEnd;

      if (startTime === undefined || endTime === undefined) {
        const fallback = this.config.fallbackDuration! / 2;
        startTime = timestamp - fallback;
        endTime = timestamp + fallback;
      }

      return this.mediaGenerator.generateAnimatedImage(
        videoPath,
        startTime,
        endTime,
        this.config.audioPadding,
        {
          fps: this.config.animatedFps,
          maxWidth: this.config.animatedMaxWidth,
          maxHeight: this.config.animatedMaxHeight,
          crf: this.config.animatedCrf,
        },
      );
    } else {
      return this.mediaGenerator.generateScreenshot(videoPath, timestamp, {
        format: this.config.imageFormat as "jpg" | "png" | "webp",
        quality: this.config.imageQuality,
        maxWidth: this.config.imageMaxWidth,
        maxHeight: this.config.imageMaxHeight,
      });
    }
  }

  private formatMiscInfoPattern(filename: string): string {
    if (!this.config.miscInfoPattern) {
      return "";
    }

    const now = new Date();
    const hours = String(now.getHours()).padStart(2, "0");
    const minutes = String(now.getMinutes()).padStart(2, "0");
    const seconds = String(now.getSeconds()).padStart(2, "0");
    const milliseconds = String(now.getMilliseconds()).padStart(3, "0");

    const filenameWithoutExt = filename.replace(/\.[^.]+$/, "");
    const filenameWithExt = filename;

    let result = this.config.miscInfoPattern
      .replace(/%f/g, filenameWithoutExt)
      .replace(/%F/g, filenameWithExt)
      .replace(/%t/g, `${hours}:${minutes}:${seconds}`)
      .replace(/%T/g, `${hours}:${minutes}:${seconds}:${milliseconds}`)
      .replace(/<br>/g, "\n");

    return result;
  }

  private generateAudioFilename(): string {
    const timestamp = Date.now();
    return `audio_${timestamp}.mp3`;
  }

  private generateImageFilename(): string {
    const timestamp = Date.now();
    const ext =
      this.config.imageType === "avif" ? "avif" : this.config.imageFormat;
    return `image_${timestamp}.${ext}`;
  }

  private showStatusNotification(message: string): void {
    const type = this.config.notificationType || "osd";

    if (type === "osd" || type === "both") {
      this.showOsdNotification(message);
    }

    if ((type === "system" || type === "both") && this.notificationCallback) {
      this.notificationCallback("SubMiner", { body: message });
    }
  }

  private showOsdNotification(text: string): void {
    if (this.osdCallback) {
      this.osdCallback(text);
    } else if (this.mpvClient && this.mpvClient.send) {
      this.mpvClient.send({
        command: ["show-text", text, "3000"],
      });
    }
  }

  private async showNotification(
    noteId: number,
    label: string | number,
    errorSuffix?: string,
  ): Promise<void> {
    const message = errorSuffix
      ? `Updated card: ${label} (${errorSuffix})`
      : `Updated card: ${label}`;

    const type = this.config.notificationType || "osd";

    if (type === "osd" || type === "both") {
      this.showOsdNotification(message);
    }

    if ((type === "system" || type === "both") && this.notificationCallback) {
      let notificationIconPath: string | undefined;

      if (this.mpvClient && this.mpvClient.currentVideoPath) {
        try {
          const timestamp = this.mpvClient.currentTimePos || 0;
          const iconBuffer = await this.mediaGenerator.generateNotificationIcon(
            this.mpvClient.currentVideoPath,
            timestamp,
          );
          if (iconBuffer && iconBuffer.length > 0) {
            notificationIconPath =
              this.mediaGenerator.writeNotificationIconToFile(
                iconBuffer,
                noteId,
              );
          }
        } catch (err) {
          console.warn(
            "Failed to generate notification icon:",
            (err as Error).message,
          );
        }
      }

      this.notificationCallback("Anki Card Updated", {
        body: message,
        icon: notificationIconPath,
      });

      if (notificationIconPath) {
        this.mediaGenerator.scheduleNotificationIconCleanup(
          notificationIconPath,
        );
      }
    }
  }

  private mergeFieldValue(
    existing: string,
    newValue: string,
    overwrite: boolean,
  ): string {
    if (overwrite || !existing.trim()) {
      return newValue;
    }
    if (this.config.mediaInsertMode === "prepend") {
      return newValue + existing;
    }
    return existing + newValue;
  }

  /**
   * Update the last added Anki card using subtitle blocks from clipboard.
   * This is the manual update flow (animecards-style) when auto-update is disabled.
   */
  async updateLastAddedFromClipboard(clipboardText: string): Promise<void> {
    try {
      if (!clipboardText || !clipboardText.trim()) {
        this.showOsdNotification("Clipboard is empty");
        return;
      }

      if (!this.mpvClient || !this.mpvClient.currentVideoPath) {
        this.showOsdNotification("No video loaded");
        return;
      }

      // Parse clipboard into blocks (separated by blank lines)
      const blocks = clipboardText
        .split(/\n\s*\n/)
        .map((b) => b.trim())
        .filter((b) => b.length > 0);

      if (blocks.length === 0) {
        this.showOsdNotification("No subtitle blocks found in clipboard");
        return;
      }

      // Lookup timings for each block
      const timings: { startTime: number; endTime: number }[] = [];
      for (const block of blocks) {
        const timing = this.timingTracker.findTiming(block);
        if (timing) {
          timings.push(timing);
        }
      }

      if (timings.length === 0) {
        this.showOsdNotification(
          "Subtitle timing not found; copy again while playing",
        );
        return;
      }

      // Compute range from all matched timings
      const rangeStart = Math.min(...timings.map((t) => t.startTime));
      let rangeEnd = Math.max(...timings.map((t) => t.endTime));

      const maxMediaDuration = this.config.maxMediaDuration ?? 30;
      if (maxMediaDuration > 0 && rangeEnd - rangeStart > maxMediaDuration) {
        console.warn(
          `Media range ${(rangeEnd - rangeStart).toFixed(1)}s exceeds cap of ${maxMediaDuration}s, clamping`,
        );
        rangeEnd = rangeStart + maxMediaDuration;
      }

      this.showOsdNotification("Updating card from clipboard...");
      this.updateInProgress = true;

      try {
        // Get last added note
        const query = this.config.deck
          ? `"deck:${this.config.deck}" added:1`
          : "added:1";
        const noteIds = (await this.client.findNotes(query)) as number[];
        if (!noteIds || noteIds.length === 0) {
          this.showOsdNotification("No recently added cards found");
          return;
        }

        // Get max note ID (most recent)
        const noteId = Math.max(...noteIds);

        // Get note info for expression
        const notesInfoResult = await this.client.notesInfo([noteId]);
        const notesInfo = notesInfoResult as unknown as NoteInfo[];
        if (!notesInfo || notesInfo.length === 0) {
          this.showOsdNotification("Card not found");
          return;
        }

        const noteInfo = notesInfo[0];
        const fields = this.extractFields(noteInfo.fields);
        const expressionText = fields.expression || fields.word || "";

        // Build sentence from blocks (join with spaces between blocks)
        const sentence = blocks.join(" ");
        const updatedFields: Record<string, string> = {};
        let updatePerformed = false;
        const errors: string[] = [];

        // Add sentence field
        if (this.config.sentenceField) {
          const processedSentence = this.processSentence(sentence, fields);
          updatedFields[this.config.sentenceField] = processedSentence;
          updatePerformed = true;
        }

        console.log(
          `Clipboard update: timing range ${rangeStart.toFixed(2)}s - ${rangeEnd.toFixed(2)}s`,
        );

        // Generate and upload audio
        if (this.config.generateAudio) {
          try {
            const audioFilename = this.generateAudioFilename();
            const audioBuffer = await this.mediaGenerator.generateAudio(
              this.mpvClient.currentVideoPath,
              rangeStart,
              rangeEnd,
              this.config.audioPadding,
            );

            if (audioBuffer) {
              await this.client.storeMediaFile(audioFilename, audioBuffer);
              const existingAudio =
                noteInfo.fields[this.config.audioField!]?.value || "";
              updatedFields[this.config.audioField!] = this.mergeFieldValue(
                existingAudio,
                `[sound:${audioFilename}]`,
                this.config.overwriteAudio !== false,
              );

              if (this.config.miscInfoField) {
                const miscInfo = this.formatMiscInfoPattern(audioFilename);
                if (miscInfo) {
                  updatedFields[this.config.miscInfoField] = miscInfo;
                }
              }

              updatePerformed = true;
            }
          } catch (error) {
            console.error(
              "Failed to generate audio:",
              (error as Error).message,
            );
            errors.push("audio");
          }
        }

        // Generate and upload image
        if (this.config.generateImage) {
          try {
            const imageFilename = this.generateImageFilename();
            let imageBuffer: Buffer | null = null;

            if (this.config.imageType === "avif") {
              imageBuffer = await this.mediaGenerator.generateAnimatedImage(
                this.mpvClient.currentVideoPath,
                rangeStart,
                rangeEnd,
                this.config.audioPadding,
                {
                  fps: this.config.animatedFps,
                  maxWidth: this.config.animatedMaxWidth,
                  maxHeight: this.config.animatedMaxHeight,
                  crf: this.config.animatedCrf,
                },
              );
            } else {
              const timestamp = this.mpvClient.currentTimePos || 0;
              imageBuffer = await this.mediaGenerator.generateScreenshot(
                this.mpvClient.currentVideoPath,
                timestamp,
                {
                  format: this.config.imageFormat as "jpg" | "png" | "webp",
                  quality: this.config.imageQuality,
                  maxWidth: this.config.imageMaxWidth,
                  maxHeight: this.config.imageMaxHeight,
                },
              );
            }

            if (imageBuffer) {
              await this.client.storeMediaFile(imageFilename, imageBuffer);
              const existingImage =
                noteInfo.fields[this.config.imageField!]?.value || "";
              updatedFields[this.config.imageField!] = this.mergeFieldValue(
                existingImage,
                `<img src="${imageFilename}">`,
                this.config.overwriteImage !== false,
              );

              if (
                this.config.miscInfoField &&
                !updatedFields[this.config.miscInfoField]
              ) {
                const miscInfo = this.formatMiscInfoPattern(imageFilename);
                if (miscInfo) {
                  updatedFields[this.config.miscInfoField] = miscInfo;
                }
              }

              updatePerformed = true;
            }
          } catch (error) {
            console.error(
              "Failed to generate image:",
              (error as Error).message,
            );
            errors.push("image");
          }
        }

        if (updatePerformed) {
          await this.client.updateNoteFields(noteId, updatedFields);
          const label = expressionText || noteId;
          console.log("Updated card from clipboard:", label);
          const errorSuffix =
            errors.length > 0 ? `${errors.join(", ")} failed` : undefined;
          await this.showNotification(noteId, label, errorSuffix);
        }
      } finally {
        this.updateInProgress = false;
      }
    } catch (error) {
      console.error(
        "Error updating card from clipboard:",
        (error as Error).message,
      );
      this.showOsdNotification(`Update failed: ${(error as Error).message}`);
    }
  }

  async markLastCardAsAudioCard(): Promise<void> {
    try {
      if (!this.mpvClient || !this.mpvClient.currentVideoPath) {
        this.showOsdNotification("No video loaded");
        return;
      }

      if (!this.mpvClient.currentSubText) {
        this.showOsdNotification("No current subtitle");
        return;
      }

      let startTime = this.mpvClient.currentSubStart;
      let endTime = this.mpvClient.currentSubEnd;

      if (startTime === undefined || endTime === undefined) {
        const currentTime = this.mpvClient.currentTimePos || 0;
        const fallback = this.config.fallbackDuration! / 2;
        startTime = currentTime - fallback;
        endTime = currentTime + fallback;
      }

      const maxMediaDuration = this.config.maxMediaDuration ?? 30;
      if (maxMediaDuration > 0 && endTime - startTime > maxMediaDuration) {
        endTime = startTime + maxMediaDuration;
      }

      this.showOsdNotification("Marking card as audio card...");
      this.updateInProgress = true;

      try {
        const query = this.config.deck
          ? `"deck:${this.config.deck}" added:1`
          : "added:1";
        const noteIds = (await this.client.findNotes(query)) as number[];
        if (!noteIds || noteIds.length === 0) {
          this.showOsdNotification("No recently added cards found");
          return;
        }

        const noteId = Math.max(...noteIds);

        const notesInfoResult = await this.client.notesInfo([noteId]);
        const notesInfo = notesInfoResult as unknown as NoteInfo[];
        if (!notesInfo || notesInfo.length === 0) {
          this.showOsdNotification("Card not found");
          return;
        }

        const noteInfo = notesInfo[0];
        const fields = this.extractFields(noteInfo.fields);
        const expressionText = fields.expression || fields.word || "";

        const updatedFields: Record<string, string> = {};
        const errors: string[] = [];

        const audioCardFieldName = this.config.audioCardField || "IsAudioCard";
        updatedFields[audioCardFieldName] = "x";

        if (this.config.sentenceField) {
          const processedSentence = this.processSentence(
            this.mpvClient.currentSubText,
            fields,
          );
          updatedFields[this.config.sentenceField] = processedSentence;
        }

        const audioFieldName =
          this.config.sentenceCardAudioField || "SentenceAudio";
        try {
          const audioFilename = this.generateAudioFilename();
          const audioBuffer = await this.mediaGenerator.generateAudio(
            this.mpvClient.currentVideoPath,
            startTime,
            endTime,
            this.config.audioPadding,
          );

          if (audioBuffer) {
            await this.client.storeMediaFile(audioFilename, audioBuffer);
            updatedFields[audioFieldName] = `[sound:${audioFilename}]`;

            if (this.config.miscInfoField) {
              const miscInfo = this.formatMiscInfoPattern(audioFilename);
              if (miscInfo) {
                updatedFields[this.config.miscInfoField] = miscInfo;
              }
            }
          }
        } catch (error) {
          console.error(
            "Failed to generate audio for audio card:",
            (error as Error).message,
          );
          errors.push("audio");
        }

        if (this.config.generateImage) {
          try {
            const imageFilename = this.generateImageFilename();
            let imageBuffer: Buffer | null = null;

            if (this.config.imageType === "avif") {
              imageBuffer = await this.mediaGenerator.generateAnimatedImage(
                this.mpvClient.currentVideoPath,
                startTime,
                endTime,
                this.config.audioPadding,
                {
                  fps: this.config.animatedFps,
                  maxWidth: this.config.animatedMaxWidth,
                  maxHeight: this.config.animatedMaxHeight,
                  crf: this.config.animatedCrf,
                },
              );
            } else {
              const timestamp = this.mpvClient.currentTimePos || 0;
              imageBuffer = await this.mediaGenerator.generateScreenshot(
                this.mpvClient.currentVideoPath,
                timestamp,
                {
                  format: this.config.imageFormat as "jpg" | "png" | "webp",
                  quality: this.config.imageQuality,
                  maxWidth: this.config.imageMaxWidth,
                  maxHeight: this.config.imageMaxHeight,
                },
              );
            }

            if (imageBuffer && this.config.imageField) {
              await this.client.storeMediaFile(imageFilename, imageBuffer);
              updatedFields[this.config.imageField] =
                `<img src="${imageFilename}">`;

              if (
                this.config.miscInfoField &&
                !updatedFields[this.config.miscInfoField]
              ) {
                const miscInfo = this.formatMiscInfoPattern(imageFilename);
                if (miscInfo) {
                  updatedFields[this.config.miscInfoField] = miscInfo;
                }
              }
            }
          } catch (error) {
            console.error(
              "Failed to generate image for audio card:",
              (error as Error).message,
            );
            errors.push("image");
          }
        }

        await this.client.updateNoteFields(noteId, updatedFields);
        const label = expressionText || noteId;
        console.log("Marked card as audio card:", label);
        const errorSuffix =
          errors.length > 0 ? `${errors.join(", ")} failed` : undefined;
        await this.showNotification(noteId, label, errorSuffix);
      } finally {
        this.updateInProgress = false;
      }
    } catch (error) {
      console.error(
        "Error marking card as audio card:",
        (error as Error).message,
      );
      this.showOsdNotification(
        `Audio card failed: ${(error as Error).message}`,
      );
    }
  }

  async createSentenceCard(
    sentence: string,
    startTime: number,
    endTime: number,
    secondarySubText?: string,
  ): Promise<void> {
    if (!this.config.sentenceCardModel) {
      this.showOsdNotification("sentenceCardModel not configured");
      return;
    }

    if (!this.mpvClient || !this.mpvClient.currentVideoPath) {
      this.showOsdNotification("No video loaded");
      return;
    }

    const maxMediaDuration = this.config.maxMediaDuration ?? 30;
    if (maxMediaDuration > 0 && endTime - startTime > maxMediaDuration) {
      console.warn(
        `Sentence card media range ${(endTime - startTime).toFixed(1)}s exceeds cap of ${maxMediaDuration}s, clamping`,
      );
      endTime = startTime + maxMediaDuration;
    }

    this.showOsdNotification("Creating sentence card...");

    const videoPath = this.mpvClient.currentVideoPath;
    const fields: Record<string, string> = {};
    const errors: string[] = [];

    const sentenceField = this.config.sentenceCardSentenceField || "Sentence";
    const audioFieldName =
      this.config.sentenceCardAudioField || "SentenceAudio";

    fields[sentenceField] = sentence;

    if (secondarySubText) {
      fields["SelectionText"] = secondarySubText;
    }

    if (this.config.isLapis || this.config.isKiku) {
      fields["IsSentenceCard"] = "x";
      fields["Expression"] = sentence;
    }

    const deck = this.config.deck || "Default";
    let noteId: number;
    try {
      noteId = await this.client.addNote(
        deck,
        this.config.sentenceCardModel,
        fields,
      );
      console.log("Created sentence card:", noteId);
      this.previousNoteIds.add(noteId);
    } catch (error) {
      console.error(
        "Failed to create sentence card:",
        (error as Error).message,
      );
      this.showOsdNotification(
        `Sentence card failed: ${(error as Error).message}`,
      );
      return;
    }

    const mediaFields: Record<string, string> = {};

    try {
      const audioFilename = this.generateAudioFilename();
      const audioBuffer = await this.mediaGenerator.generateAudio(
        videoPath,
        startTime,
        endTime,
        this.config.audioPadding,
      );

      if (audioBuffer) {
        await this.client.storeMediaFile(audioFilename, audioBuffer);
        mediaFields[audioFieldName] = `[sound:${audioFilename}]`;
      }
    } catch (error) {
      console.error(
        "Failed to generate sentence audio:",
        (error as Error).message,
      );
      errors.push("audio");
    }

    try {
      const imageFilename = this.generateImageFilename();
      let imageBuffer: Buffer | null = null;

      if (this.config.imageType === "avif") {
        imageBuffer = await this.mediaGenerator.generateAnimatedImage(
          videoPath,
          startTime,
          endTime,
          this.config.audioPadding,
          {
            fps: this.config.animatedFps,
            maxWidth: this.config.animatedMaxWidth,
            maxHeight: this.config.animatedMaxHeight,
            crf: this.config.animatedCrf,
          },
        );
      } else {
        const timestamp = this.mpvClient.currentTimePos || 0;
        imageBuffer = await this.mediaGenerator.generateScreenshot(
          videoPath,
          timestamp,
          {
            format: this.config.imageFormat as "jpg" | "png" | "webp",
            quality: this.config.imageQuality,
            maxWidth: this.config.imageMaxWidth,
            maxHeight: this.config.imageMaxHeight,
          },
        );
      }

      if (imageBuffer && this.config.imageField) {
        await this.client.storeMediaFile(imageFilename, imageBuffer);
        mediaFields[this.config.imageField] = `<img src="${imageFilename}">`;
      }
    } catch (error) {
      console.error(
        "Failed to generate sentence image:",
        (error as Error).message,
      );
      errors.push("image");
    }

    if (Object.keys(mediaFields).length > 0) {
      try {
        await this.client.updateNoteFields(noteId, mediaFields);
      } catch (error) {
        console.error(
          "Failed to update sentence card media:",
          (error as Error).message,
        );
        errors.push("media update");
      }
    }

    const label =
      sentence.length > 30 ? sentence.substring(0, 30) + "..." : sentence;
    const errorSuffix =
      errors.length > 0 ? `${errors.join(", ")} failed` : undefined;
    await this.showNotification(noteId, label, errorSuffix);
  }

  private async findDuplicateNote(
    expression: string,
    excludeNoteId: number,
    noteInfo: NoteInfo,
  ): Promise<number | null> {
    let fieldName = "";
    for (const name of Object.keys(noteInfo.fields)) {
      if (
        ["word", "expression"].includes(name.toLowerCase()) &&
        noteInfo.fields[name].value
      ) {
        fieldName = name;
        break;
      }
    }
    if (!fieldName) return null;

    const escapedExpression = expression.replace(/"/g, '\\"');
    const deckPrefix = this.config.deck ? `"deck:${this.config.deck}" ` : "";
    const query = `${deckPrefix}"${fieldName}:${escapedExpression}"`;

    try {
      const noteIds = await this.client.findNotes(query);
      const duplicates = noteIds.filter((id) => id !== excludeNoteId);
      return duplicates.length > 0 ? duplicates[0] : null;
    } catch (error) {
      console.warn("Duplicate search failed:", (error as Error).message);
      return null;
    }
  }

  private getGroupableFieldNames(): string[] {
    const fields: string[] = [];
    if (this.config.imageField) fields.push(this.config.imageField);
    if (this.config.sentenceField) fields.push(this.config.sentenceField);
    if (this.config.audioField) fields.push(this.config.audioField);
    const sentenceAudioField =
      this.config.sentenceCardAudioField || "SentenceAudio";
    if (!fields.includes(sentenceAudioField)) fields.push(sentenceAudioField);
    if (this.config.miscInfoField) fields.push(this.config.miscInfoField);
    fields.push("SentenceFurigana");
    return fields;
  }

  private applyFieldGrouping(
    existingValue: string,
    newValue: string,
    groupId: number,
    isPictureField: boolean,
  ): string {
    if (!existingValue.trim()) return newValue;
    if (!newValue.trim()) return existingValue;

    if (isPictureField) {
      const grouped = existingValue.replace(
        /<img(?![^>]*data-group-id)([^>]*)>/g,
        `<img data-group-id="${groupId}"$1>`,
      );
      return grouped + "\n" + newValue;
    }

    const hasGroups = /data-group-id/.test(existingValue);

    if (!hasGroups) {
      return (
        `<span data-group-id="${groupId}">${existingValue}</span>\n` + newValue
      );
    }

    const groupedSpanRegex = /<span\s+data-group-id="[^"]*">[\s\S]*?<\/span>/g;
    let lastEnd = 0;
    let result = "";
    let match;

    while ((match = groupedSpanRegex.exec(existingValue)) !== null) {
      const before = existingValue.slice(lastEnd, match.index);
      if (before.trim()) {
        result += `<span data-group-id="${groupId}">${before.trim()}</span>\n`;
      }
      result += match[0] + "\n";
      lastEnd = match.index + match[0].length;
    }

    const after = existingValue.slice(lastEnd);
    if (after.trim()) {
      result += `\n<span data-group-id="${groupId}">${after.trim()}</span>`;
    }

    return result + "\n" + newValue;
  }

  private async generateMediaForMerge(): Promise<{
    audioField?: string;
    audioValue?: string;
    imageField?: string;
    imageValue?: string;
    miscInfoValue?: string;
  }> {
    const result: {
      audioField?: string;
      audioValue?: string;
      imageField?: string;
      imageValue?: string;
      miscInfoValue?: string;
    } = {};

    if (this.config.generateAudio && this.mpvClient?.currentVideoPath) {
      try {
        const audioFilename = this.generateAudioFilename();
        const audioBuffer = await this.generateAudio();
        if (audioBuffer) {
          await this.client.storeMediaFile(audioFilename, audioBuffer);
          result.audioField = this.config.audioField!;
          result.audioValue = `[sound:${audioFilename}]`;
          if (this.config.miscInfoField) {
            result.miscInfoValue = this.formatMiscInfoPattern(audioFilename);
          }
        }
      } catch (error) {
        console.error(
          "Failed to generate audio for merge:",
          (error as Error).message,
        );
      }
    }

    if (this.config.generateImage && this.mpvClient?.currentVideoPath) {
      try {
        const imageFilename = this.generateImageFilename();
        const imageBuffer = await this.generateImage();
        if (imageBuffer) {
          await this.client.storeMediaFile(imageFilename, imageBuffer);
          result.imageField = this.config.imageField!;
          result.imageValue = `<img src="${imageFilename}">`;
          if (this.config.miscInfoField && !result.miscInfoValue) {
            result.miscInfoValue = this.formatMiscInfoPattern(imageFilename);
          }
        }
      } catch (error) {
        console.error(
          "Failed to generate image for merge:",
          (error as Error).message,
        );
      }
    }

    return result;
  }

  private async performFieldGroupingMerge(
    keepNoteId: number,
    deleteNoteId: number,
    deleteNoteInfo: NoteInfo,
    expression: string,
  ): Promise<void> {
    const keepNotesInfoResult = await this.client.notesInfo([keepNoteId]);
    const keepNotesInfo = keepNotesInfoResult as unknown as NoteInfo[];
    if (!keepNotesInfo || keepNotesInfo.length === 0) {
      console.warn("Keep note not found:", keepNoteId);
      return;
    }
    const keepNoteInfo = keepNotesInfo[0];

    const media = await this.generateMediaForMerge();

    const newFields: Record<string, string> = {};
    if (this.config.sentenceField && this.mpvClient.currentSubText) {
      const deleteFields = this.extractFields(deleteNoteInfo.fields);
      newFields[this.config.sentenceField] = this.processSentence(
        this.mpvClient.currentSubText,
        deleteFields,
      );
    }
    if (media.audioField && media.audioValue) {
      newFields[media.audioField] = media.audioValue;
    }
    if (media.imageField && media.imageValue) {
      newFields[media.imageField] = media.imageValue;
    }
    if (this.config.miscInfoField && media.miscInfoValue) {
      newFields[this.config.miscInfoField] = media.miscInfoValue;
    }

    const groupableFields = this.getGroupableFieldNames();
    const mergedFields: Record<string, string> = {};

    for (const fieldName of groupableFields) {
      const existingValue = keepNoteInfo.fields[fieldName]?.value || "";
      const newValue = newFields[fieldName] || "";
      const isPictureField = fieldName === this.config.imageField;

      if (existingValue.trim() && newValue.trim()) {
        mergedFields[fieldName] = this.applyFieldGrouping(
          existingValue,
          newValue,
          keepNoteId,
          isPictureField,
        );
      } else if (newValue.trim()) {
        mergedFields[fieldName] = newValue;
      }
    }

    if (Object.keys(mergedFields).length > 0) {
      await this.client.updateNoteFields(keepNoteId, mergedFields);
    }

    await this.client.deleteNotes([deleteNoteId]);
    this.previousNoteIds.delete(deleteNoteId);

    console.log("Merged duplicate card:", expression, "into note:", keepNoteId);
    this.showStatusNotification(`Merged duplicate: ${expression}`);
    await this.showNotification(keepNoteId, expression);
  }

  private async handleFieldGroupingAuto(
    originalNoteId: number,
    newNoteId: number,
    newNoteInfo: NoteInfo,
    expression: string,
  ): Promise<void> {
    try {
      await this.performFieldGroupingMerge(
        originalNoteId,
        newNoteId,
        newNoteInfo,
        expression,
      );
    } catch (error) {
      console.error(
        "Field grouping auto merge failed:",
        (error as Error).message,
      );
      this.showOsdNotification(
        `Field grouping failed: ${(error as Error).message}`,
      );
    }
  }

  private async handleFieldGroupingManual(
    originalNoteId: number,
    newNoteId: number,
    newNoteInfo: NoteInfo,
    expression: string,
  ): Promise<boolean> {
    if (!this.fieldGroupingCallback) {
      console.warn(
        "No field grouping callback registered, skipping manual mode",
      );
      return false;
    }

    try {
      const originalNotesInfoResult = await this.client.notesInfo([
        originalNoteId,
      ]);
      const originalNotesInfo =
        originalNotesInfoResult as unknown as NoteInfo[];
      if (!originalNotesInfo || originalNotesInfo.length === 0) {
        return false;
      }
      const originalNoteInfo = originalNotesInfo[0];

      const originalFields = this.extractFields(originalNoteInfo.fields);
      const newFields = this.extractFields(newNoteInfo.fields);

      const originalCard: KikuDuplicateCardInfo = {
        noteId: originalNoteId,
        expression:
          originalFields.expression || originalFields.word || expression,
        sentencePreview: this.truncateSentence(
          originalFields[
            (this.config.sentenceField || "sentence").toLowerCase()
          ] || "",
        ),
        hasAudio: !!(
          originalNoteInfo.fields[this.config.audioField!]?.value ||
          originalNoteInfo.fields[
            this.config.sentenceCardAudioField || "SentenceAudio"
          ]?.value
        ),
        hasImage: !!originalNoteInfo.fields[this.config.imageField!]?.value,
        isOriginal: true,
      };

      const newCard: KikuDuplicateCardInfo = {
        noteId: newNoteId,
        expression: newFields.expression || newFields.word || expression,
        sentencePreview: this.truncateSentence(
          newFields[(this.config.sentenceField || "sentence").toLowerCase()] ||
            this.mpvClient.currentSubText ||
            "",
        ),
        hasAudio: false,
        hasImage: false,
        isOriginal: false,
      };

      const choice = await this.fieldGroupingCallback({
        original: originalCard,
        duplicate: newCard,
      });

      if (choice.cancelled) {
        return false;
      }

      const keepNoteId = choice.keepNoteId;
      const deleteNoteId = choice.deleteNoteId;
      const deleteNoteInfo =
        deleteNoteId === newNoteId ? newNoteInfo : originalNoteInfo;

      await this.performFieldGroupingMerge(
        keepNoteId,
        deleteNoteId,
        deleteNoteInfo,
        expression,
      );
      return true;
    } catch (error) {
      console.error(
        "Field grouping manual merge failed:",
        (error as Error).message,
      );
      this.showOsdNotification(
        `Field grouping failed: ${(error as Error).message}`,
      );
      return false;
    }
  }

  private truncateSentence(sentence: string): string {
    const clean = sentence.replace(/<[^>]*>/g, "").trim();
    if (clean.length <= 100) return clean;
    return clean.substring(0, 100) + "...";
  }

  destroy(): void {
    this.stop();
    this.mediaGenerator.cleanup();
  }
}
