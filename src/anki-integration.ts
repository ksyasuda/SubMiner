/*
 * mpv-yomitan - Yomitan integration for mpv
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
import { AnkiConnectConfig } from "./types";

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
  private mpvClient: any;
  private osdCallback: ((text: string) => void) | null = null;
  private notificationCallback: ((title: string, options: any) => void) | null = null;
  private updateInProgress = false;

  constructor(
    config: AnkiConnectConfig,
    timingTracker: SubtitleTimingTracker,
    mpvClient: any,
    osdCallback?: (text: string) => void,
    notificationCallback?: (title: string, options: any) => void,
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
      miscInfoPattern: "[mpv-yomitan] %f (%t)",
      notificationType: "osd",
      imageQuality: 92,
      animatedFps: 10,
      animatedMaxWidth: 640,
      animatedCrf: 35,
      autoUpdateNewCards: true,
      maxMediaDuration: 30,
      ...config,
    };

    this.client = new AnkiConnectClient(this.config.url!);
    this.mediaGenerator = new MediaGenerator();
    this.timingTracker = timingTracker;
    this.mpvClient = mpvClient;
    this.osdCallback = osdCallback || null;
    this.notificationCallback = notificationCallback || null;
  }

  start(): void {
    if (this.pollingInterval) {
      this.stop();
    }

    console.log("Starting AnkiConnect integration with polling rate:", this.config.pollingRate);
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

    try {
      const query = this.config.deck
        ? `"deck:${this.config.deck}" added:1`
        : "added:1";
      const noteIds = (await this.client.findNotes(query, { maxRetries: 0 })) as number[];
      const currentNoteIds = new Set(noteIds);

      if (!this.initialized) {
        this.previousNoteIds = currentNoteIds;
        this.initialized = true;
        console.log(`AnkiConnect initialized with ${currentNoteIds.size} existing cards`);
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
    }
  }

  private async processNewCard(noteId: number): Promise<void> {
    this.updateInProgress = true;
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
            const existingAudio = noteInfo.fields[this.config.audioField!]?.value || "";
            updatedFields[this.config.audioField!] = this.mergeFieldValue(
              existingAudio, `[sound:${audioFilename}]`, this.config.overwriteAudio !== false,
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
          this.showOsdNotification(`Audio generation failed: ${(error as Error).message}`);
        }
      }

      let imageBuffer: Buffer | null = null;
      if (this.config.generateImage && this.mpvClient) {
        try {
          const imageFilename = this.generateImageFilename();
          imageBuffer = await this.generateImage();

          if (imageBuffer) {
            await this.client.storeMediaFile(imageFilename, imageBuffer);
            const existingImage = noteInfo.fields[this.config.imageField!]?.value || "";
            updatedFields[this.config.imageField!] = this.mergeFieldValue(
              existingImage, `<img src="${imageFilename}">`, this.config.overwriteImage !== false,
            );

            if (this.config.miscInfoField && !updatedFields[this.config.miscInfoField]) {
              const miscInfo = this.formatMiscInfoPattern(imageFilename);
              if (miscInfo) {
                updatedFields[this.config.miscInfoField] = miscInfo;
              }
            }

            updatePerformed = true;
          }
        } catch (error) {
          console.error("Failed to generate image:", (error as Error).message);
          this.showOsdNotification(`Image generation failed: ${(error as Error).message}`);
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
    } finally {
      this.updateInProgress = false;
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

    const sentenceFieldName = this.config.sentenceField?.toLowerCase() || 'sentence';
    const existingSentence = noteFields[sentenceFieldName] || '';

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
    if (!this.mpvClient || !this.mpvClient.currentVideoPath) {
      return null;
    }

    const videoPath = this.mpvClient.currentVideoPath;
    let startTime = this.mpvClient.currentSubStart;
    let endTime = this.mpvClient.currentSubEnd;

    if (startTime === undefined || endTime === undefined) {
      const currentTime = this.mpvClient.currentTimePos || 0;
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
      return this.mediaGenerator.generateScreenshot(
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
    const ext = this.config.imageType === "avif" ? "avif" : this.config.imageFormat;
    return `image_${timestamp}.${ext}`;
  }

  private showStatusNotification(message: string): void {
    const type = this.config.notificationType || "osd";

    if (type === "osd" || type === "both") {
      this.showOsdNotification(message);
    }

    if ((type === "system" || type === "both") && this.notificationCallback) {
      this.notificationCallback("mpv-yomitan", { body: message });
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

  private async showNotification(noteId: number, label: string | number, errorSuffix?: string): Promise<void> {
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
            notificationIconPath = this.mediaGenerator.writeNotificationIconToFile(
              iconBuffer,
              noteId,
            );
          }
        } catch (err) {
          console.warn('Failed to generate notification icon:', (err as Error).message);
        }
      }

      this.notificationCallback('Anki Card Updated', {
        body: message,
        icon: notificationIconPath,
      });
    }
  }

  private mergeFieldValue(existing: string, newValue: string, overwrite: boolean): string {
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

      console.log(`Clipboard update: timing range ${rangeStart.toFixed(2)}s - ${rangeEnd.toFixed(2)}s`);

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
            const existingAudio = noteInfo.fields[this.config.audioField!]?.value || "";
            updatedFields[this.config.audioField!] = this.mergeFieldValue(
              existingAudio, `[sound:${audioFilename}]`, this.config.overwriteAudio !== false,
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
            const existingImage = noteInfo.fields[this.config.imageField!]?.value || "";
            updatedFields[this.config.imageField!] = this.mergeFieldValue(
              existingImage, `<img src="${imageFilename}">`, this.config.overwriteImage !== false,
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
          errors.push("image");
        }
      }

      if (updatePerformed) {
        await this.client.updateNoteFields(noteId, updatedFields);
        const label = expressionText || noteId;
        console.log("Updated card from clipboard:", label);
        const errorSuffix = errors.length > 0 ? `${errors.join(", ")} failed` : undefined;
        await this.showNotification(noteId, label, errorSuffix);
      }
      } finally {
        this.updateInProgress = false;
      }
    } catch (error) {
      console.error("Error updating card from clipboard:", (error as Error).message);
      this.showOsdNotification(
        `Update failed: ${(error as Error).message}`,
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
    const audioFieldName = this.config.sentenceCardAudioField || "SentenceAudio";

    fields[sentenceField] = sentence;

    if (secondarySubText) {
      fields["SelectionText"] = secondarySubText;
    }

    if (this.config.isLapis) {
      fields["IsSentenceCard"] = "x";
      fields["Expression"] = sentence;
    }

    const deck = this.config.deck || "Default";
    let noteId: number;
    try {
      noteId = await this.client.addNote(deck, this.config.sentenceCardModel, fields);
      console.log("Created sentence card:", noteId);
      this.previousNoteIds.add(noteId);
    } catch (error) {
      console.error("Failed to create sentence card:", (error as Error).message);
      this.showOsdNotification(`Sentence card failed: ${(error as Error).message}`);
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
      console.error("Failed to generate sentence audio:", (error as Error).message);
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
      console.error("Failed to generate sentence image:", (error as Error).message);
      errors.push("image");
    }

    if (Object.keys(mediaFields).length > 0) {
      try {
        await this.client.updateNoteFields(noteId, mediaFields);
      } catch (error) {
        console.error("Failed to update sentence card media:", (error as Error).message);
        errors.push("media update");
      }
    }

    const label = sentence.length > 30 ? sentence.substring(0, 30) + "..." : sentence;
    const errorSuffix = errors.length > 0 ? `${errors.join(", ")} failed` : undefined;
    await this.showNotification(noteId, label, errorSuffix);
  }

  destroy(): void {
    this.stop();
    this.mediaGenerator.cleanup();
  }
}
