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
  private config: AnkiConnectConfig;
  private pollingInterval: ReturnType<typeof setInterval> | null = null;
  private previousNoteIds = new Set<number>();
  private initialized = false;
  private backoffMs = 200;
  private maxBackoffMs = 5000;
  private mpvClient: any;
  private osdCallback: ((text: string) => void) | null = null;
  private notificationCallback: ((title: string, options: any) => void) | null = null;

  constructor(
    config: AnkiConnectConfig,
    _timingTracker: SubtitleTimingTracker,
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
      audioPadding: 0.5,
      fallbackDuration: 3.0,
      miscInfoPattern: "[mpv-yomitan] %f (%t)",
      showNotificationOnUpdate: false,
      imageQuality: 92,
      animatedFps: 10,
      animatedMaxWidth: 640,
      animatedCrf: 35,
      ...config,
    };

    this.client = new AnkiConnectClient(this.config.url!);
    this.mediaGenerator = new MediaGenerator();
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
    try {
      const query = this.config.deck
        ? `"deck:${this.config.deck}" added:1`
        : "added:1";
      const noteIds = (await this.client.findNotes(query)) as number[];
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

        for (const noteId of newNoteIds) {
          await this.processNewCard(noteId);
        }
      }

      this.backoffMs = 200;
    } catch (error) {
      this.backoffMs = Math.min(this.backoffMs * 2, this.maxBackoffMs);
      if (this.backoffMs > 200) {
        console.warn("AnkiConnect polling failed, backing off...");
      }
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
            updatedFields[this.config.audioField!] = `[sound:${audioFilename}]`;

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
            updatedFields[this.config.imageField!] = `<img src="${imageFilename}">`;

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
        this.showOsdNotification(`Updated card: ${expressionText}`);

        if (this.config.showNotificationOnUpdate && this.notificationCallback) {
          // Generate a separate notification icon as a static PNG.
          // This is independent of the Picture field (which may be AVIF).
          // Using a temp file path is more reliable on Linux/Wayland.
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
              // Continue without icon - notifications are best-effort
            }
          }

          this.notificationCallback('Anki Card Updated', {
            body: `Updated card: ${expressionText}`,
            icon: notificationIconPath,
          });
        }
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

  private showOsdNotification(text: string): void {
    if (this.osdCallback) {
      this.osdCallback(text);
    } else if (this.mpvClient && this.mpvClient.send) {
      this.mpvClient.send({
        command: ["show-text", text, "3000"],
      });
    }
  }

  destroy(): void {
    this.stop();
    this.mediaGenerator.cleanup();
  }
}
