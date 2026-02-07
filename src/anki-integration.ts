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
import * as path from "path";
import {
  AnkiConnectConfig,
  KikuDuplicateCardInfo,
  KikuFieldGroupingChoice,
  KikuMergePreviewResponse,
  MpvClient,
  NotificationOptions,
} from "./types";

interface NoteInfo {
  noteId: number;
  fields: Record<string, { value: string }>;
}

type CardKind = "sentence" | "audio";

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
  private progressDepth = 0;
  private progressTimer: ReturnType<typeof setInterval> | null = null;
  private progressMessage = "";
  private progressFrame = 0;
  private parseWarningKeys = new Set<string>();
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
      miscInfoField: "MiscInfo",
      miscInfoPattern: "[SubMiner] %f (%t)",
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
    this.fieldGroupingCallback = fieldGroupingCallback || null;
  }

  private getLapisConfig(): {
    enabled: boolean;
    sentenceCardModel?: string;
    sentenceCardSentenceField?: string;
    sentenceCardAudioField?: string;
  } {
    const lapis = this.config.isLapis;
    return {
      enabled: lapis?.enabled === true,
      sentenceCardModel: lapis?.sentenceCardModel,
      sentenceCardSentenceField: lapis?.sentenceCardSentenceField,
      sentenceCardAudioField: lapis?.sentenceCardAudioField,
    };
  }

  private getKikuConfig(): {
    enabled: boolean;
    fieldGrouping?: "auto" | "manual" | "disabled";
    deleteDuplicateInAuto?: boolean;
  } {
    const kiku = this.config.isKiku;
    return {
      enabled: kiku?.enabled === true,
      fieldGrouping: kiku?.fieldGrouping,
      deleteDuplicateInAuto: kiku?.deleteDuplicateInAuto,
    };
  }

  private getEffectiveSentenceCardConfig(): {
    model?: string;
    sentenceField: string;
    audioField: string;
    lapisEnabled: boolean;
    kikuEnabled: boolean;
    kikuFieldGrouping: "auto" | "manual" | "disabled";
    kikuDeleteDuplicateInAuto: boolean;
  } {
    const lapis = this.getLapisConfig();
    const kiku = this.getKikuConfig();

    return {
      model: lapis.sentenceCardModel,
      sentenceField: lapis.sentenceCardSentenceField || "Sentence",
      audioField: lapis.sentenceCardAudioField || "SentenceAudio",
      lapisEnabled: lapis.enabled,
      kikuEnabled: kiku.enabled,
      kikuFieldGrouping: (kiku.fieldGrouping || "disabled") as
        | "auto"
        | "manual"
        | "disabled",
      kikuDeleteDuplicateInAuto: kiku.deleteDuplicateInAuto !== false,
    };
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

  private async processNewCard(
    noteId: number,
    options?: { skipKikuFieldGrouping?: boolean },
  ): Promise<void> {
    this.beginUpdateProgress("Updating card");
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

      const sentenceCardConfig = this.getEffectiveSentenceCardConfig();
      if (
        !options?.skipKikuFieldGrouping &&
        sentenceCardConfig.kikuEnabled &&
        sentenceCardConfig.kikuFieldGrouping !== "disabled"
      ) {
        const duplicateNoteId = await this.findDuplicateNote(
          expressionText,
          noteId,
          noteInfo,
        );
        if (duplicateNoteId !== null) {
          if (sentenceCardConfig.kikuFieldGrouping === "auto") {
            await this.handleFieldGroupingAuto(
              duplicateNoteId,
              noteId,
              noteInfo,
              expressionText,
            );
            return;
          } else if (sentenceCardConfig.kikuFieldGrouping === "manual") {
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
      let miscInfoFilename: string | null = null;
      const sentenceField = sentenceCardConfig.sentenceField;

      if (sentenceField && this.mpvClient.currentSubText) {
        const processedSentence = this.processSentence(
          this.mpvClient.currentSubText,
          fields,
        );
        updatedFields[sentenceField] = processedSentence;
        updatePerformed = true;
      }

      if (this.config.generateAudio && this.mpvClient) {
        try {
          const audioFilename = this.generateAudioFilename();
          const audioBuffer = await this.generateAudio();

          if (audioBuffer) {
            await this.client.storeMediaFile(audioFilename, audioBuffer);
            const sentenceAudioField =
              this.getResolvedSentenceAudioFieldName(noteInfo);
            if (sentenceAudioField) {
              const existingAudio =
                noteInfo.fields[sentenceAudioField]?.value || "";
              updatedFields[sentenceAudioField] = this.mergeFieldValue(
                existingAudio,
                `[sound:${audioFilename}]`,
                this.config.overwriteAudio !== false,
              );
            }
            miscInfoFilename = audioFilename;
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
            miscInfoFilename = imageFilename;
            updatePerformed = true;
          }
        } catch (error) {
          console.error("Failed to generate image:", (error as Error).message);
          this.showOsdNotification(
            `Image generation failed: ${(error as Error).message}`,
          );
        }
      }

      if (this.config.miscInfoField) {
        const miscInfo = this.formatMiscInfoPattern(
          miscInfoFilename || "",
          this.mpvClient.currentSubStart,
        );
        const miscInfoField = this.getResolvedConfiguredFieldName(
          noteInfo,
          this.config.miscInfoField,
        );
        if (miscInfo && miscInfoField) {
          updatedFields[miscInfoField] = miscInfo;
          updatePerformed = true;
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
      this.endUpdateProgress();
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
      this.mpvClient.currentAudioStreamIndex,
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

  private formatMiscInfoPattern(
    fallbackFilename: string,
    startTimeSeconds?: number,
  ): string {
    if (!this.config.miscInfoPattern) {
      return "";
    }

    const currentVideoPath = this.mpvClient.currentVideoPath || "";
    const videoFilename = currentVideoPath
      ? path.basename(currentVideoPath)
      : "";
    const filenameWithExt = videoFilename || fallbackFilename;
    const filenameWithoutExt = filenameWithExt.replace(/\.[^.]+$/, "");

    const currentTimePos =
      typeof startTimeSeconds === "number" && Number.isFinite(startTimeSeconds)
        ? startTimeSeconds
        : this.mpvClient.currentTimePos;
    let totalMilliseconds = 0;
    if (Number.isFinite(currentTimePos) && currentTimePos >= 0) {
      totalMilliseconds = Math.floor(currentTimePos * 1000);
    } else {
      const now = new Date();
      totalMilliseconds =
        now.getHours() * 3600000 +
        now.getMinutes() * 60000 +
        now.getSeconds() * 1000 +
        now.getMilliseconds();
    }

    const totalSeconds = Math.floor(totalMilliseconds / 1000);
    const hours = String(Math.floor(totalSeconds / 3600)).padStart(2, "0");
    const minutes = String(Math.floor((totalSeconds % 3600) / 60)).padStart(
      2,
      "0",
    );
    const seconds = String(totalSeconds % 60).padStart(2, "0");
    const milliseconds = String(totalMilliseconds % 1000).padStart(3, "0");

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

  private beginUpdateProgress(initialMessage: string): void {
    this.progressDepth += 1;
    if (this.progressDepth > 1) return;

    this.progressMessage = initialMessage;
    this.progressFrame = 0;
    this.showProgressTick();
    this.progressTimer = setInterval(() => {
      this.showProgressTick();
    }, 180);
  }

  private endUpdateProgress(): void {
    this.progressDepth = Math.max(0, this.progressDepth - 1);
    if (this.progressDepth > 0) return;

    if (this.progressTimer) {
      clearInterval(this.progressTimer);
      this.progressTimer = null;
    }
    this.progressMessage = "";
    this.progressFrame = 0;
  }

  private showProgressTick(): void {
    if (!this.progressMessage) return;
    const frames = ["|", "/", "-", "\\"];
    const frame = frames[this.progressFrame % frames.length];
    this.progressFrame += 1;
    this.showOsdNotification(`${this.progressMessage} ${frame}`);
  }

  private async withUpdateProgress<T>(
    initialMessage: string,
    action: () => Promise<T>,
  ): Promise<T> {
    this.beginUpdateProgress(initialMessage);
    this.updateInProgress = true;
    try {
      return await action();
    } finally {
      this.updateInProgress = false;
      this.endUpdateProgress();
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

  private resolveFieldName(
    availableFieldNames: string[],
    preferredName: string,
  ): string | null {
    const exact = availableFieldNames.find((name) => name === preferredName);
    if (exact) return exact;

    const lower = preferredName.toLowerCase();
    const ci = availableFieldNames.find((name) => name.toLowerCase() === lower);
    return ci || null;
  }

  private resolveNoteFieldName(
    noteInfo: NoteInfo,
    preferredName?: string,
  ): string | null {
    if (!preferredName) return null;
    return this.resolveFieldName(Object.keys(noteInfo.fields), preferredName);
  }

  private resolveFirstNoteFieldName(
    noteInfo: NoteInfo,
    preferredNames: (string | undefined)[],
  ): string | null {
    for (const preferredName of preferredNames) {
      const resolved = this.resolveNoteFieldName(noteInfo, preferredName);
      if (resolved) return resolved;
    }
    return null;
  }

  private warnFieldParseOnce(
    fieldName: string,
    reason: string,
    detail?: string,
  ): void {
    const key = `${fieldName.toLowerCase()}::${reason}`;
    if (this.parseWarningKeys.has(key)) return;
    this.parseWarningKeys.add(key);
    const suffix = detail ? ` (${detail})` : "";
    console.warn(
      `Field grouping parse warning [${fieldName}] ${reason}${suffix}`,
    );
  }

  private setCardTypeFields(
    updatedFields: Record<string, string>,
    availableFieldNames: string[],
    cardKind: CardKind,
  ): void {
    const audioFlagNames = ["IsAudioCard"];

    if (cardKind === "sentence") {
      const sentenceFlag = this.resolveFieldName(
        availableFieldNames,
        "IsSentenceCard",
      );
      if (sentenceFlag) {
        updatedFields[sentenceFlag] = "x";
      }

      for (const audioFlagName of audioFlagNames) {
        const resolved = this.resolveFieldName(
          availableFieldNames,
          audioFlagName,
        );
        if (resolved && resolved !== sentenceFlag) {
          updatedFields[resolved] = "";
        }
      }

      const wordAndSentenceFlag = this.resolveFieldName(
        availableFieldNames,
        "IsWordAndSentenceCard",
      );
      if (wordAndSentenceFlag && wordAndSentenceFlag !== sentenceFlag) {
        updatedFields[wordAndSentenceFlag] = "";
      }
      return;
    }

    const resolvedAudioFlags = Array.from(
      new Set(
        audioFlagNames
          .map((name) => this.resolveFieldName(availableFieldNames, name))
          .filter((name): name is string => Boolean(name)),
      ),
    );
    const audioFlagName = resolvedAudioFlags[0] || null;
    if (audioFlagName) {
      updatedFields[audioFlagName] = "x";
    }
    for (const extraAudioFlag of resolvedAudioFlags.slice(1)) {
      updatedFields[extraAudioFlag] = "";
    }

    const sentenceFlag = this.resolveFieldName(
      availableFieldNames,
      "IsSentenceCard",
    );
    if (sentenceFlag && sentenceFlag !== audioFlagName) {
      updatedFields[sentenceFlag] = "";
    }

    const wordAndSentenceFlag = this.resolveFieldName(
      availableFieldNames,
      "IsWordAndSentenceCard",
    );
    if (wordAndSentenceFlag && wordAndSentenceFlag !== audioFlagName) {
      updatedFields[wordAndSentenceFlag] = "";
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
      this.beginUpdateProgress("Updating card from clipboard");
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
        const sentenceAudioField =
          this.getResolvedSentenceAudioFieldName(noteInfo);
        const sentenceField =
          this.getEffectiveSentenceCardConfig().sentenceField;

        // Build sentence from blocks (join with spaces between blocks)
        const sentence = blocks.join(" ");
        const updatedFields: Record<string, string> = {};
        let updatePerformed = false;
        const errors: string[] = [];
        let miscInfoFilename: string | null = null;

        // Add sentence field
        if (sentenceField) {
          const processedSentence = this.processSentence(sentence, fields);
          updatedFields[sentenceField] = processedSentence;
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
              this.mpvClient.currentAudioStreamIndex,
            );

            if (audioBuffer) {
              await this.client.storeMediaFile(audioFilename, audioBuffer);
              if (sentenceAudioField) {
                const existingAudio =
                  noteInfo.fields[sentenceAudioField]?.value || "";
                updatedFields[sentenceAudioField] = this.mergeFieldValue(
                  existingAudio,
                  `[sound:${audioFilename}]`,
                  this.config.overwriteAudio !== false,
                );
              }
              miscInfoFilename = audioFilename;
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
              miscInfoFilename = imageFilename;
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

        if (this.config.miscInfoField) {
          const miscInfo = this.formatMiscInfoPattern(
            miscInfoFilename || "",
            rangeStart,
          );
          const miscInfoField = this.getResolvedConfiguredFieldName(
            noteInfo,
            this.config.miscInfoField,
          );
          if (miscInfo && miscInfoField) {
            updatedFields[miscInfoField] = miscInfo;
            updatePerformed = true;
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
        this.endUpdateProgress();
      }
    } catch (error) {
      console.error(
        "Error updating card from clipboard:",
        (error as Error).message,
      );
      this.showOsdNotification(`Update failed: ${(error as Error).message}`);
    }
  }

  async triggerFieldGroupingForLastAddedCard(): Promise<void> {
    const sentenceCardConfig = this.getEffectiveSentenceCardConfig();
    if (!sentenceCardConfig.kikuEnabled) {
      this.showOsdNotification("Kiku mode is not enabled");
      return;
    }
    if (sentenceCardConfig.kikuFieldGrouping === "disabled") {
      this.showOsdNotification("Kiku field grouping is disabled");
      return;
    }

    if (this.updateInProgress) {
      this.showOsdNotification("Anki update already in progress");
      return;
    }

    try {
      await this.withUpdateProgress("Grouping duplicate cards", async () => {
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
        const noteInfoBeforeUpdate = notesInfo[0];

        // Run the update pass first when required media fields are missing.
        if (
          !this.hasAllConfiguredFields(noteInfoBeforeUpdate, [
            this.config.imageField,
          ])
        ) {
          await this.processNewCard(noteId, { skipKikuFieldGrouping: true });
        }

        const refreshedInfoResult = await this.client.notesInfo([noteId]);
        const refreshedInfo = refreshedInfoResult as unknown as NoteInfo[];
        if (!refreshedInfo || refreshedInfo.length === 0) {
          this.showOsdNotification("Card not found");
          return;
        }

        const noteInfo = refreshedInfo[0];
        const fields = this.extractFields(noteInfo.fields);
        const expressionText = fields.expression || fields.word || "";
        if (!expressionText) {
          this.showOsdNotification("No expression/word field found");
          return;
        }

        const duplicateNoteId = await this.findDuplicateNote(
          expressionText,
          noteId,
          noteInfo,
        );
        if (duplicateNoteId === null) {
          this.showOsdNotification("No duplicate card found");
          return;
        }

        if (sentenceCardConfig.kikuFieldGrouping === "auto") {
          await this.handleFieldGroupingAuto(
            duplicateNoteId,
            noteId,
            noteInfo,
            expressionText,
          );
          return;
        }
        const handled = await this.handleFieldGroupingManual(
          duplicateNoteId,
          noteId,
          noteInfo,
          expressionText,
        );
        if (!handled) {
          this.showOsdNotification("Field grouping cancelled");
        }
      });
    } catch (error) {
      console.error(
        "Error triggering field grouping:",
        (error as Error).message,
      );
      this.showOsdNotification(
        `Field grouping failed: ${(error as Error).message}`,
      );
    }
  }

  async markLastCardAsAudioCard(): Promise<void> {
    if (this.updateInProgress) {
      this.showOsdNotification("Anki update already in progress");
      return;
    }

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
      await this.withUpdateProgress("Marking audio card", async () => {
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
        let miscInfoFilename: string | null = null;

        this.setCardTypeFields(
          updatedFields,
          Object.keys(noteInfo.fields),
          "audio",
        );

        if (this.config.sentenceField) {
          const processedSentence = this.processSentence(
            this.mpvClient.currentSubText,
            fields,
          );
          updatedFields[this.config.sentenceField] = processedSentence;
        }

        const sentenceCardConfig = this.getEffectiveSentenceCardConfig();
        const audioFieldName = sentenceCardConfig.audioField;
        try {
          const audioFilename = this.generateAudioFilename();
          const audioBuffer = await this.mediaGenerator.generateAudio(
            this.mpvClient.currentVideoPath,
            startTime,
            endTime,
            this.config.audioPadding,
            this.mpvClient.currentAudioStreamIndex,
          );

          if (audioBuffer) {
            await this.client.storeMediaFile(audioFilename, audioBuffer);
            updatedFields[audioFieldName] = `[sound:${audioFilename}]`;
            miscInfoFilename = audioFilename;
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
              miscInfoFilename = imageFilename;
            }
          } catch (error) {
            console.error(
              "Failed to generate image for audio card:",
              (error as Error).message,
            );
            errors.push("image");
          }
        }

        if (this.config.miscInfoField) {
          const miscInfo = this.formatMiscInfoPattern(
            miscInfoFilename || "",
            startTime,
          );
          const miscInfoField = this.getResolvedConfiguredFieldName(
            noteInfo,
            this.config.miscInfoField,
          );
          if (miscInfo && miscInfoField) {
            updatedFields[miscInfoField] = miscInfo;
          }
        }

        await this.client.updateNoteFields(noteId, updatedFields);
        const label = expressionText || noteId;
        console.log("Marked card as audio card:", label);
        const errorSuffix =
          errors.length > 0 ? `${errors.join(", ")} failed` : undefined;
        await this.showNotification(noteId, label, errorSuffix);
      });
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
    if (this.updateInProgress) {
      this.showOsdNotification("Anki update already in progress");
      return;
    }

    const sentenceCardConfig = this.getEffectiveSentenceCardConfig();
    const sentenceCardModel = sentenceCardConfig.model;
    if (!sentenceCardModel) {
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
    await this.withUpdateProgress("Creating sentence card", async () => {
      const videoPath = this.mpvClient.currentVideoPath;
      const fields: Record<string, string> = {};
      const errors: string[] = [];
      let miscInfoFilename: string | null = null;

      const sentenceField = sentenceCardConfig.sentenceField;
      const audioFieldName = sentenceCardConfig.audioField || "SentenceAudio";
      let resolvedMiscInfoField: string | null = null;
      let resolvedSentenceAudioField: string = audioFieldName;
      let resolvedExpressionAudioField: string | null = null;

      fields[sentenceField] = sentence;

      if (secondarySubText) {
        fields["SelectionText"] = secondarySubText;
      }

      if (sentenceCardConfig.lapisEnabled || sentenceCardConfig.kikuEnabled) {
        fields["IsSentenceCard"] = "x";
        fields["Expression"] = sentence;
      }

      const deck = this.config.deck || "Default";
      let noteId: number;
      try {
        noteId = await this.client.addNote(
          deck,
          sentenceCardModel,
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

      try {
        const noteInfoResult = await this.client.notesInfo([noteId]);
        const noteInfos = noteInfoResult as unknown as NoteInfo[];
        if (noteInfos.length > 0) {
          const createdNoteInfo = noteInfos[0];
          resolvedSentenceAudioField =
            this.resolveNoteFieldName(createdNoteInfo, audioFieldName) ||
            audioFieldName;
          resolvedExpressionAudioField = this.resolveFirstNoteFieldName(
            createdNoteInfo,
            [this.config.audioField || "ExpressionAudio", "ExpressionAudio"],
          );
          resolvedMiscInfoField = this.getResolvedConfiguredFieldName(
            createdNoteInfo,
            this.config.miscInfoField,
          );
          const cardTypeFields: Record<string, string> = {};
          this.setCardTypeFields(
            cardTypeFields,
            Object.keys(createdNoteInfo.fields),
            "sentence",
          );
          if (Object.keys(cardTypeFields).length > 0) {
            await this.client.updateNoteFields(noteId, cardTypeFields);
          }
        }
      } catch (error) {
        console.error(
          "Failed to normalize sentence card type fields:",
          (error as Error).message,
        );
        errors.push("card type fields");
      }

      const mediaFields: Record<string, string> = {};

      try {
        const audioFilename = this.generateAudioFilename();
        const audioBuffer = await this.mediaGenerator.generateAudio(
          videoPath,
          startTime,
          endTime,
          this.config.audioPadding,
          this.mpvClient.currentAudioStreamIndex,
        );

        if (audioBuffer) {
          await this.client.storeMediaFile(audioFilename, audioBuffer);
          const audioValue = `[sound:${audioFilename}]`;
          mediaFields[resolvedSentenceAudioField] = audioValue;
          if (
            resolvedExpressionAudioField &&
            resolvedExpressionAudioField !== resolvedSentenceAudioField
          ) {
            mediaFields[resolvedExpressionAudioField] = audioValue;
          }
          miscInfoFilename = audioFilename;
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
          miscInfoFilename = imageFilename;
        }
      } catch (error) {
        console.error(
          "Failed to generate sentence image:",
          (error as Error).message,
        );
        errors.push("image");
      }

      if (this.config.miscInfoField) {
        const miscInfo = this.formatMiscInfoPattern(
          miscInfoFilename || "",
          startTime,
        );
        if (miscInfo && resolvedMiscInfoField) {
          mediaFields[resolvedMiscInfoField] = miscInfo;
        }
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
    });
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

    const escapedFieldName = this.escapeAnkiSearchValue(fieldName);
    const escapedExpression = this.escapeAnkiSearchValue(expression);
    const deckPrefix = this.config.deck ? `"deck:${this.config.deck}" ` : "";
    const query = `${deckPrefix}"${escapedFieldName}:${escapedExpression}"`;

    try {
      const noteIds = (await this.client.findNotes(query)) as number[];
      return await this.findFirstExactDuplicateNoteId(
        noteIds,
        excludeNoteId,
        fieldName,
        expression,
      );
    } catch (error) {
      console.warn("Duplicate search failed:", (error as Error).message);
      return null;
    }
  }

  private escapeAnkiSearchValue(value: string): string {
    return value.replace(/([\\":*?()[\]{}])/g, "\\$1");
  }

  private normalizeDuplicateValue(value: string): string {
    return value.replace(/\s+/g, " ").trim();
  }

  private async findFirstExactDuplicateNoteId(
    candidateNoteIds: number[],
    excludeNoteId: number,
    fieldName: string,
    expression: string,
  ): Promise<number | null> {
    const candidates = candidateNoteIds.filter((id) => id !== excludeNoteId);
    if (candidates.length === 0) return null;

    const normalizedExpression = this.normalizeDuplicateValue(expression);
    const chunkSize = 50;
    for (let i = 0; i < candidates.length; i += chunkSize) {
      const chunk = candidates.slice(i, i + chunkSize);
      const notesInfoResult = await this.client.notesInfo(chunk);
      const notesInfo = notesInfoResult as unknown as NoteInfo[];
      for (const noteInfo of notesInfo) {
        const resolvedField = this.resolveNoteFieldName(noteInfo, fieldName);
        if (!resolvedField) continue;
        const candidateValue = noteInfo.fields[resolvedField]?.value || "";
        if (
          this.normalizeDuplicateValue(candidateValue) === normalizedExpression
        ) {
          return noteInfo.noteId;
        }
      }
    }

    return null;
  }

  private getGroupableFieldNames(): string[] {
    const fields: string[] = [];
    fields.push("Sentence");
    fields.push("SentenceAudio");
    fields.push("Picture");
    if (this.config.imageField) fields.push(this.config.imageField);
    if (this.config.sentenceField) fields.push(this.config.sentenceField);
    if (
      this.config.audioField &&
      this.config.audioField.toLowerCase() !== "expressionaudio"
    ) {
      fields.push(this.config.audioField);
    }
    const sentenceCardConfig = this.getEffectiveSentenceCardConfig();
    const sentenceAudioField = sentenceCardConfig.audioField;
    if (!fields.includes(sentenceAudioField)) fields.push(sentenceAudioField);
    if (this.config.miscInfoField) fields.push(this.config.miscInfoField);
    fields.push("SentenceFurigana");
    return fields;
  }

  private getPreferredSentenceAudioFieldName(): string {
    const sentenceCardConfig = this.getEffectiveSentenceCardConfig();
    return sentenceCardConfig.audioField || "SentenceAudio";
  }

  private getResolvedSentenceAudioFieldName(noteInfo: NoteInfo): string | null {
    return (
      this.resolveNoteFieldName(
        noteInfo,
        this.getPreferredSentenceAudioFieldName(),
      ) || this.getResolvedConfiguredFieldName(noteInfo, this.config.audioField)
    );
  }

  private extractUngroupedValue(value: string): string {
    const groupedSpanRegex = /<span\s+data-group-id="[^"]*">[\s\S]*?<\/span>/gi;
    const ungrouped = value.replace(groupedSpanRegex, "").trim();
    if (ungrouped) return ungrouped;
    return value.trim();
  }

  private extractLastSoundTag(value: string): string {
    const matches = value.match(/\[sound:[^\]]+\]/g);
    if (!matches || matches.length === 0) return "";
    return matches[matches.length - 1];
  }

  private extractLastImageTag(value: string): string {
    const matches = value.match(/<img\b[^>]*>/gi);
    if (!matches || matches.length === 0) return "";
    return matches[matches.length - 1];
  }

  private extractImageTags(value: string): string[] {
    const matches = value.match(/<img\b[^>]*>/gi);
    return matches || [];
  }

  private ensureImageGroupId(imageTag: string, groupId: number): string {
    if (!imageTag) return "";
    if (/data-group-id=/i.test(imageTag)) {
      return imageTag.replace(
        /data-group-id="[^"]*"/i,
        `data-group-id="${groupId}"`,
      );
    }
    return imageTag.replace(/<img\b/i, `<img data-group-id="${groupId}"`);
  }

  private extractSpanEntries(
    value: string,
    fieldName: string,
  ): { groupId: number; content: string }[] {
    const entries: { groupId: number; content: string }[] = [];
    const malformedIdRegex = /<span\s+[^>]*data-group-id="([^"]*)"[^>]*>/gi;
    let malformed;
    while ((malformed = malformedIdRegex.exec(value)) !== null) {
      const rawId = malformed[1];
      const groupId = Number(rawId);
      if (!Number.isFinite(groupId) || groupId <= 0) {
        this.warnFieldParseOnce(fieldName, "invalid-group-id", rawId);
      }
    }

    const spanRegex = /<span\s+data-group-id="(\d+)"[^>]*>([\s\S]*?)<\/span>/gi;
    let match;
    while ((match = spanRegex.exec(value)) !== null) {
      const groupId = Number(match[1]);
      if (!Number.isFinite(groupId) || groupId <= 0) continue;
      const content = this.normalizeStrictGroupedValue(
        match[2] || "",
        fieldName,
      );
      if (!content) {
        this.warnFieldParseOnce(fieldName, "empty-group-content");
        continue;
      }
      entries.push({ groupId, content });
    }
    if (entries.length === 0 && /<span\b/i.test(value)) {
      this.warnFieldParseOnce(fieldName, "no-usable-span-entries");
    }
    return entries;
  }

  private parseStrictEntries(
    value: string,
    fallbackGroupId: number,
    fieldName: string,
  ): { groupId: number; content: string }[] {
    const entries = this.extractSpanEntries(value, fieldName);
    if (entries.length === 0) {
      const ungrouped = this.normalizeStrictGroupedValue(
        this.extractUngroupedValue(value),
        fieldName,
      );
      if (ungrouped) {
        entries.push({ groupId: fallbackGroupId, content: ungrouped });
      }
    }

    const unique: { groupId: number; content: string }[] = [];
    const seen = new Set<string>();
    for (const entry of entries) {
      const key = `${entry.groupId}::${entry.content}`;
      if (seen.has(key)) continue;
      seen.add(key);
      unique.push(entry);
    }
    return unique;
  }

  private parsePictureEntries(
    value: string,
    fallbackGroupId: number,
  ): { groupId: number; tag: string }[] {
    const tags = this.extractImageTags(value);
    const result: { groupId: number; tag: string }[] = [];
    for (const tag of tags) {
      const idMatch = tag.match(/data-group-id="(\d+)"/i);
      let groupId = fallbackGroupId;
      if (idMatch) {
        const parsed = Number(idMatch[1]);
        if (!Number.isFinite(parsed) || parsed <= 0) {
          this.warnFieldParseOnce("Picture", "invalid-group-id", idMatch[1]);
        } else {
          groupId = parsed;
        }
      }
      const normalizedTag = this.ensureImageGroupId(tag, groupId);
      if (!normalizedTag) {
        this.warnFieldParseOnce("Picture", "empty-image-tag");
        continue;
      }
      result.push({ groupId, tag: normalizedTag });
    }
    return result;
  }

  private normalizeStrictGroupedValue(
    value: string,
    fieldName: string,
  ): string {
    const ungrouped = this.extractUngroupedValue(value);
    if (!ungrouped) return "";

    const normalizedField = fieldName.toLowerCase();
    if (
      normalizedField === "sentenceaudio" ||
      normalizedField === "expressionaudio"
    ) {
      const lastSoundTag = this.extractLastSoundTag(ungrouped);
      if (!lastSoundTag) {
        this.warnFieldParseOnce(fieldName, "missing-sound-tag");
      }
      return lastSoundTag || ungrouped;
    }

    if (normalizedField === "picture") {
      const lastImageTag = this.extractLastImageTag(ungrouped);
      if (!lastImageTag) {
        this.warnFieldParseOnce(fieldName, "missing-image-tag");
      }
      return lastImageTag || ungrouped;
    }

    return ungrouped;
  }

  private shouldUseStrictSpanGrouping(fieldName: string): boolean {
    const normalized = fieldName.toLowerCase();
    const configuredMiscInfo = (this.config.miscInfoField || "").toLowerCase();
    if (
      [
        "picture",
        "sentence",
        "sentenceaudio",
        "sentencefurigana",
        "miscinfo",
      ].includes(normalized)
    ) {
      return true;
    }
    if (configuredMiscInfo && normalized === configuredMiscInfo) {
      return true;
    }
    return false;
  }

  private applyFieldGrouping(
    existingValue: string,
    newValue: string,
    keepGroupId: number,
    sourceGroupId: number,
    fieldName: string,
  ): string {
    if (this.shouldUseStrictSpanGrouping(fieldName)) {
      if (fieldName.toLowerCase() === "picture") {
        const keepEntries = this.parsePictureEntries(
          existingValue,
          keepGroupId,
        );
        const sourceEntries = this.parsePictureEntries(newValue, sourceGroupId);
        if (keepEntries.length === 0 && sourceEntries.length === 0) {
          return existingValue || newValue;
        }
        const mergedTags = keepEntries.map((entry) =>
          this.ensureImageGroupId(entry.tag, entry.groupId),
        );
        const seen = new Set(mergedTags);
        for (const entry of sourceEntries) {
          const normalized = this.ensureImageGroupId(entry.tag, entry.groupId);
          if (seen.has(normalized)) continue;
          seen.add(normalized);
          mergedTags.push(normalized);
        }
        return mergedTags.join("");
      }

      const keepEntries = this.parseStrictEntries(
        existingValue,
        keepGroupId,
        fieldName,
      );
      const sourceEntries = this.parseStrictEntries(
        newValue,
        sourceGroupId,
        fieldName,
      );
      if (keepEntries.length === 0 && sourceEntries.length === 0) {
        return existingValue || newValue;
      }
      if (sourceEntries.length === 0) {
        return keepEntries
          .map(
            (entry) =>
              `<span data-group-id="${entry.groupId}">${entry.content}</span>`,
          )
          .join("");
      }
      const merged = [...keepEntries];
      const seen = new Set(
        keepEntries.map((entry) => `${entry.groupId}::${entry.content}`),
      );
      for (const entry of sourceEntries) {
        const key = `${entry.groupId}::${entry.content}`;
        if (seen.has(key)) continue;
        seen.add(key);
        merged.push(entry);
      }
      if (merged.length === 0) return existingValue;
      return merged
        .map(
          (entry) =>
            `<span data-group-id="${entry.groupId}">${entry.content}</span>`,
        )
        .join("");
    }

    if (!existingValue.trim()) return newValue;
    if (!newValue.trim()) return existingValue;

    const hasGroups = /data-group-id/.test(existingValue);

    if (!hasGroups) {
      return (
        `<span data-group-id="${keepGroupId}">${existingValue}</span>\n` +
        newValue
      );
    }

    const groupedSpanRegex = /<span\s+data-group-id="[^"]*">[\s\S]*?<\/span>/g;
    let lastEnd = 0;
    let result = "";
    let match;

    while ((match = groupedSpanRegex.exec(existingValue)) !== null) {
      const before = existingValue.slice(lastEnd, match.index);
      if (before.trim()) {
        result += `<span data-group-id="${keepGroupId}">${before.trim()}</span>\n`;
      }
      result += match[0] + "\n";
      lastEnd = match.index + match[0].length;
    }

    const after = existingValue.slice(lastEnd);
    if (after.trim()) {
      result += `\n<span data-group-id="${keepGroupId}">${after.trim()}</span>`;
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
          result.audioField = this.getPreferredSentenceAudioFieldName();
          result.audioValue = `[sound:${audioFilename}]`;
          if (this.config.miscInfoField) {
            result.miscInfoValue = this.formatMiscInfoPattern(
              audioFilename,
              this.mpvClient.currentSubStart,
            );
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
            result.miscInfoValue = this.formatMiscInfoPattern(
              imageFilename,
              this.mpvClient.currentSubStart,
            );
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

  private getResolvedFieldValue(
    noteInfo: NoteInfo,
    preferredFieldName?: string,
  ): string {
    if (!preferredFieldName) return "";
    const resolved = this.resolveNoteFieldName(noteInfo, preferredFieldName);
    if (!resolved) return "";
    return noteInfo.fields[resolved]?.value || "";
  }

  private getResolvedConfiguredFieldName(
    noteInfo: NoteInfo,
    preferredFieldName?: string,
  ): string | null {
    return this.resolveNoteFieldName(noteInfo, preferredFieldName);
  }

  private async computeFieldGroupingMergedFields(
    keepNoteId: number,
    deleteNoteId: number,
    keepNoteInfo: NoteInfo,
    deleteNoteInfo: NoteInfo,
    includeGeneratedMedia: boolean,
  ): Promise<Record<string, string>> {
    const groupableFields = this.getGroupableFieldNames();
    const keepFieldNames = Object.keys(keepNoteInfo.fields);
    const sourceFields: Record<string, string> = {};
    const resolvedKeepFieldByPreferred = new Map<string, string>();
    for (const preferredFieldName of groupableFields) {
      sourceFields[preferredFieldName] = this.getResolvedFieldValue(
        deleteNoteInfo,
        preferredFieldName,
      );
      const keepResolved = this.resolveFieldName(
        keepFieldNames,
        preferredFieldName,
      );
      if (keepResolved) {
        resolvedKeepFieldByPreferred.set(preferredFieldName, keepResolved);
      }
    }

    if (!sourceFields["SentenceFurigana"] && sourceFields["Sentence"]) {
      sourceFields["SentenceFurigana"] = sourceFields["Sentence"];
    }
    if (!sourceFields["Sentence"] && sourceFields["SentenceFurigana"]) {
      sourceFields["Sentence"] = sourceFields["SentenceFurigana"];
    }
    if (!sourceFields["Expression"] && sourceFields["Word"]) {
      sourceFields["Expression"] = sourceFields["Word"];
    }
    if (!sourceFields["Word"] && sourceFields["Expression"]) {
      sourceFields["Word"] = sourceFields["Expression"];
    }
    if (!sourceFields["SentenceAudio"] && sourceFields["ExpressionAudio"]) {
      sourceFields["SentenceAudio"] = sourceFields["ExpressionAudio"];
    }
    if (!sourceFields["ExpressionAudio"] && sourceFields["SentenceAudio"]) {
      sourceFields["ExpressionAudio"] = sourceFields["SentenceAudio"];
    }

    if (
      this.config.sentenceField &&
      !sourceFields[this.config.sentenceField] &&
      this.mpvClient.currentSubText
    ) {
      const deleteFields = this.extractFields(deleteNoteInfo.fields);
      sourceFields[this.config.sentenceField] = this.processSentence(
        this.mpvClient.currentSubText,
        deleteFields,
      );
    }

    if (includeGeneratedMedia) {
      const media = await this.generateMediaForMerge();
      if (
        media.audioField &&
        media.audioValue &&
        !sourceFields[media.audioField]
      ) {
        sourceFields[media.audioField] = media.audioValue;
      }
      if (
        media.imageField &&
        media.imageValue &&
        !sourceFields[media.imageField]
      ) {
        sourceFields[media.imageField] = media.imageValue;
      }
      if (
        this.config.miscInfoField &&
        media.miscInfoValue &&
        !sourceFields[this.config.miscInfoField]
      ) {
        sourceFields[this.config.miscInfoField] = media.miscInfoValue;
      }
    }

    const mergedFields: Record<string, string> = {};
    for (const preferredFieldName of groupableFields) {
      const keepFieldName =
        resolvedKeepFieldByPreferred.get(preferredFieldName);
      if (!keepFieldName) continue;

      const keepFieldNormalized = keepFieldName.toLowerCase();
      if (
        keepFieldNormalized === "expression" ||
        keepFieldNormalized === "expressionfurigana" ||
        keepFieldNormalized === "expressionreading" ||
        keepFieldNormalized === "expressionaudio"
      ) {
        continue;
      }

      const existingValue = keepNoteInfo.fields[keepFieldName]?.value || "";
      const newValue = sourceFields[preferredFieldName] || "";
      const isStrictField = this.shouldUseStrictSpanGrouping(keepFieldName);
      if (!existingValue.trim() && !newValue.trim()) continue;

      if (isStrictField) {
        mergedFields[keepFieldName] = this.applyFieldGrouping(
          existingValue,
          newValue,
          keepNoteId,
          deleteNoteId,
          keepFieldName,
        );
      } else if (existingValue.trim() && newValue.trim()) {
        mergedFields[keepFieldName] = this.applyFieldGrouping(
          existingValue,
          newValue,
          keepNoteId,
          deleteNoteId,
          keepFieldName,
        );
      } else {
        if (!newValue.trim()) continue;
        mergedFields[keepFieldName] = newValue;
      }
    }

    return mergedFields;
  }

  private getNoteFieldMap(noteInfo: NoteInfo): Record<string, string> {
    const fields: Record<string, string> = {};
    for (const [name, field] of Object.entries(noteInfo.fields)) {
      fields[name] = field?.value || "";
    }
    return fields;
  }

  async buildFieldGroupingPreview(
    keepNoteId: number,
    deleteNoteId: number,
    deleteDuplicate: boolean,
  ): Promise<KikuMergePreviewResponse> {
    try {
      const notesInfoResult = await this.client.notesInfo([
        keepNoteId,
        deleteNoteId,
      ]);
      const notesInfo = notesInfoResult as unknown as NoteInfo[];
      const keepNoteInfo = notesInfo.find((note) => note.noteId === keepNoteId);
      const deleteNoteInfo = notesInfo.find(
        (note) => note.noteId === deleteNoteId,
      );

      if (!keepNoteInfo || !deleteNoteInfo) {
        return { ok: false, error: "Could not load selected notes" };
      }

      const mergedFields = await this.computeFieldGroupingMergedFields(
        keepNoteId,
        deleteNoteId,
        keepNoteInfo,
        deleteNoteInfo,
        false,
      );
      const keepBefore = this.getNoteFieldMap(keepNoteInfo);
      const keepAfter = { ...keepBefore, ...mergedFields };
      const sourceBefore = this.getNoteFieldMap(deleteNoteInfo);

      const compactFields: Record<string, string> = {};
      for (const fieldName of [
        "Sentence",
        "SentenceFurigana",
        "SentenceAudio",
        "Picture",
        "MiscInfo",
      ]) {
        const resolved = this.resolveFieldName(
          Object.keys(keepAfter),
          fieldName,
        );
        if (!resolved) continue;
        compactFields[fieldName] = keepAfter[resolved] || "";
      }

      return {
        ok: true,
        compact: {
          action: {
            keepNoteId,
            deleteNoteId,
            deleteDuplicate,
          },
          mergedFields: compactFields,
        },
        full: {
          keepNote: {
            id: keepNoteId,
            fieldsBefore: keepBefore,
          },
          sourceNote: {
            id: deleteNoteId,
            fieldsBefore: sourceBefore,
          },
          result: {
            fieldsAfter: keepAfter,
            wouldDeleteNoteId: deleteDuplicate ? deleteNoteId : null,
          },
        },
      };
    } catch (error) {
      return {
        ok: false,
        error: `Failed to build preview: ${(error as Error).message}`,
      };
    }
  }

  private async performFieldGroupingMerge(
    keepNoteId: number,
    deleteNoteId: number,
    deleteNoteInfo: NoteInfo,
    expression: string,
    deleteDuplicate = true,
  ): Promise<void> {
    const keepNotesInfoResult = await this.client.notesInfo([keepNoteId]);
    const keepNotesInfo = keepNotesInfoResult as unknown as NoteInfo[];
    if (!keepNotesInfo || keepNotesInfo.length === 0) {
      console.warn("Keep note not found:", keepNoteId);
      return;
    }
    const keepNoteInfo = keepNotesInfo[0];
    const mergedFields = await this.computeFieldGroupingMergedFields(
      keepNoteId,
      deleteNoteId,
      keepNoteInfo,
      deleteNoteInfo,
      true,
    );

    if (Object.keys(mergedFields).length > 0) {
      await this.client.updateNoteFields(keepNoteId, mergedFields);
    }

    if (deleteDuplicate) {
      await this.client.deleteNotes([deleteNoteId]);
      this.previousNoteIds.delete(deleteNoteId);
    }

    console.log("Merged duplicate card:", expression, "into note:", keepNoteId);
    this.showStatusNotification(
      deleteDuplicate
        ? `Merged duplicate: ${expression}`
        : `Grouped duplicate (kept both): ${expression}`,
    );
    await this.showNotification(keepNoteId, expression);
  }

  private async handleFieldGroupingAuto(
    originalNoteId: number,
    newNoteId: number,
    newNoteInfo: NoteInfo,
    expression: string,
  ): Promise<void> {
    try {
      const sentenceCardConfig = this.getEffectiveSentenceCardConfig();
      await this.performFieldGroupingMerge(
        originalNoteId,
        newNoteId,
        newNoteInfo,
        expression,
        sentenceCardConfig.kikuDeleteDuplicateInAuto,
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
      const sentenceCardConfig = this.getEffectiveSentenceCardConfig();

      const originalFields = this.extractFields(originalNoteInfo.fields);
      const newFields = this.extractFields(newNoteInfo.fields);

      const originalCard: KikuDuplicateCardInfo = {
        noteId: originalNoteId,
        expression:
          originalFields.expression || originalFields.word || expression,
        sentencePreview: this.truncateSentence(
          originalFields[
            (sentenceCardConfig.sentenceField || "sentence").toLowerCase()
          ] || "",
        ),
        hasAudio:
          this.hasFieldValue(originalNoteInfo, this.config.audioField) ||
          this.hasFieldValue(originalNoteInfo, sentenceCardConfig.audioField),
        hasImage: this.hasFieldValue(originalNoteInfo, this.config.imageField),
        isOriginal: true,
      };

      const newCard: KikuDuplicateCardInfo = {
        noteId: newNoteId,
        expression: newFields.expression || newFields.word || expression,
        sentencePreview: this.truncateSentence(
          newFields[
            (sentenceCardConfig.sentenceField || "sentence").toLowerCase()
          ] ||
            this.mpvClient.currentSubText ||
            "",
        ),
        hasAudio:
          this.hasFieldValue(newNoteInfo, this.config.audioField) ||
          this.hasFieldValue(newNoteInfo, sentenceCardConfig.audioField),
        hasImage: this.hasFieldValue(newNoteInfo, this.config.imageField),
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
        choice.deleteDuplicate,
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

  private hasFieldValue(
    noteInfo: NoteInfo,
    preferredFieldName?: string,
  ): boolean {
    const resolved = this.resolveNoteFieldName(noteInfo, preferredFieldName);
    if (!resolved) return false;
    return Boolean(noteInfo.fields[resolved]?.value);
  }

  private hasAllConfiguredFields(
    noteInfo: NoteInfo,
    configuredFieldNames: (string | undefined)[],
  ): boolean {
    const requiredFields = configuredFieldNames.filter(
      (fieldName): fieldName is string => Boolean(fieldName),
    );
    if (requiredFields.length === 0) return true;
    return requiredFields.every((fieldName) =>
      this.hasFieldValue(noteInfo, fieldName),
    );
  }

  private async refreshMiscInfoField(
    noteId: number,
    noteInfo: NoteInfo,
  ): Promise<void> {
    if (!this.config.miscInfoField || !this.config.miscInfoPattern) return;

    const resolvedMiscField = this.resolveNoteFieldName(
      noteInfo,
      this.config.miscInfoField,
    );
    if (!resolvedMiscField) return;

    const nextValue = this.formatMiscInfoPattern(
      "",
      this.mpvClient.currentSubStart,
    );
    if (!nextValue) return;

    const currentValue = noteInfo.fields[resolvedMiscField]?.value || "";
    if (currentValue === nextValue) return;

    await this.client.updateNoteFields(noteId, {
      [resolvedMiscField]: nextValue,
    });
  }


  destroy(): void {
    this.stop();
    this.mediaGenerator.cleanup();
  }
}
