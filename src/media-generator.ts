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

import { ExecFileException, execFile } from "child_process";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";

export class MediaGenerator {
  private tempDir: string;
  private notifyIconDir: string;

  constructor(tempDir?: string) {
    this.tempDir = tempDir || path.join(os.tmpdir(), "subminer-media");
    this.notifyIconDir = path.join(os.tmpdir(), "subminer-notify");
    if (!fs.existsSync(this.tempDir)) {
      fs.mkdirSync(this.tempDir, { recursive: true });
    }
    if (!fs.existsSync(this.notifyIconDir)) {
      fs.mkdirSync(this.notifyIconDir, { recursive: true });
    }
    // Clean up old notification icons on startup (older than 1 hour)
    this.cleanupOldNotificationIcons();
  }

  /**
   * Clean up notification icons older than 1 hour.
   * Called on startup to prevent accumulation of temp files.
   */
  private cleanupOldNotificationIcons(): void {
    try {
      if (!fs.existsSync(this.notifyIconDir)) return;

      const files = fs.readdirSync(this.notifyIconDir);
      const oneHourAgo = Date.now() - 60 * 60 * 1000;

      for (const file of files) {
        if (!file.endsWith(".png")) continue;
        const filePath = path.join(this.notifyIconDir, file);
        try {
          const stat = fs.statSync(filePath);
          if (stat.mtimeMs < oneHourAgo) {
            fs.unlinkSync(filePath);
          }
        } catch (err) {
          console.debug(
            `Failed to clean up ${filePath}:`,
            (err as Error).message,
          );
        }
      }
    } catch (err) {
      console.error("Failed to cleanup old notification icons:", err);
    }
  }

  /**
   * Write a notification icon buffer to a temp file and return the file path.
   * The file path can be passed directly to Electron Notification for better
   * compatibility with Linux/Wayland notification daemons.
   */
  writeNotificationIconToFile(iconBuffer: Buffer, noteId: number): string {
    const filename = `icon_${noteId}_${Date.now()}.png`;
    const filePath = path.join(this.notifyIconDir, filename);
    fs.writeFileSync(filePath, iconBuffer);
    return filePath;
  }

  scheduleNotificationIconCleanup(filePath: string, delayMs = 10000): void {
    setTimeout(() => {
      try {
        fs.unlinkSync(filePath);
      } catch {}
    }, delayMs);
  }

  private ffmpegError(label: string, error: ExecFileException): Error {
    if (error.code === "ENOENT") {
      return new Error(
        "FFmpeg not found. Install FFmpeg to enable media generation.",
      );
    }
    return new Error(`FFmpeg ${label} failed: ${error.message}`);
  }

  async generateAudio(
    videoPath: string,
    startTime: number,
    endTime: number,
    padding: number = 0.5,
    audioStreamIndex: number | null = null,
  ): Promise<Buffer> {
    const start = Math.max(0, startTime - padding);
    const duration = endTime - startTime + 2 * padding;

    return new Promise((resolve, reject) => {
      const outputPath = path.join(this.tempDir, `audio_${Date.now()}.mp3`);
      const args: string[] = [
        "-ss",
        start.toString(),
        "-t",
        duration.toString(),
        "-i",
        videoPath,
      ];

      if (
        typeof audioStreamIndex === "number" &&
        Number.isInteger(audioStreamIndex) &&
        audioStreamIndex >= 0
      ) {
        args.push("-map", `0:${audioStreamIndex}`);
      }

      args.push(
        "-vn",
        "-acodec",
        "libmp3lame",
        "-q:a",
        "2",
        "-ar",
        "44100",
        "-y",
        outputPath,
      );

      execFile("ffmpeg", args, { timeout: 30000 }, (error) => {
        if (error) {
          reject(this.ffmpegError("audio generation", error));
          return;
        }

        try {
          const data = fs.readFileSync(outputPath);
          fs.unlinkSync(outputPath);
          resolve(data);
        } catch (err) {
          reject(err);
        }
      });
    });
  }

  async generateScreenshot(
    videoPath: string,
    timestamp: number,
    options: {
      format: "jpg" | "png" | "webp";
      quality?: number;
      maxWidth?: number;
      maxHeight?: number;
    },
  ): Promise<Buffer> {
    const { format, quality = 92, maxWidth, maxHeight } = options;
    const ext = format === "webp" ? "webp" : format === "png" ? "png" : "jpg";
    const codecMap: Record<string, string> = {
      jpg: "mjpeg",
      png: "png",
      webp: "webp",
    };

    const args: string[] = [
      "-ss",
      timestamp.toString(),
      "-i",
      videoPath,
      "-vframes",
      "1",
    ];

    const vfParts: string[] = [];
    if (maxWidth && maxWidth > 0 && maxHeight && maxHeight > 0) {
      vfParts.push(
        `scale=w=${maxWidth}:h=${maxHeight}:force_original_aspect_ratio=decrease`,
      );
    } else if (maxWidth && maxWidth > 0) {
      vfParts.push(`scale=w=${maxWidth}:h=-2`);
    } else if (maxHeight && maxHeight > 0) {
      vfParts.push(`scale=w=-2:h=${maxHeight}`);
    }
    if (vfParts.length > 0) {
      args.push("-vf", vfParts.join(","));
    }

    args.push("-c:v", codecMap[format]);

    if (format !== "png") {
      const clampedQuality = Math.max(1, Math.min(100, quality));
      if (format === "jpg") {
        const qv = Math.round(2 + (100 - clampedQuality) * (29 / 99));
        args.push("-q:v", qv.toString());
      } else {
        args.push("-q:v", clampedQuality.toString());
      }
    }

    args.push("-y");

    return new Promise((resolve, reject) => {
      const outputPath = path.join(
        this.tempDir,
        `screenshot_${Date.now()}.${ext}`,
      );
      args.push(outputPath);

      execFile("ffmpeg", args, { timeout: 30000 }, (error) => {
        if (error) {
          reject(this.ffmpegError("screenshot generation", error));
          return;
        }

        try {
          const data = fs.readFileSync(outputPath);
          fs.unlinkSync(outputPath);
          resolve(data);
        } catch (err) {
          reject(err);
        }
      });
    });
  }

  /**
   * Generate a small PNG icon suitable for desktop notifications.
   * Always outputs PNG format (known-good for Electron + Linux notification daemons).
   * Scaled to 256px width for fast encoding and small file size.
   */
  async generateNotificationIcon(
    videoPath: string,
    timestamp: number,
  ): Promise<Buffer> {
    return new Promise((resolve, reject) => {
      const outputPath = path.join(
        this.tempDir,
        `notify_icon_${Date.now()}.png`,
      );

      execFile(
        "ffmpeg",
        [
          "-ss",
          timestamp.toString(),
          "-i",
          videoPath,
          "-vframes",
          "1",
          "-vf",
          "scale=256:256:force_original_aspect_ratio=decrease,pad=256:256:(ow-iw)/2:(oh-ih)/2:black",
          "-c:v",
          "png",
          "-y",
          outputPath,
        ],
        { timeout: 30000 },
        (error) => {
          if (error) {
            reject(this.ffmpegError("notification icon generation", error));
            return;
          }

          try {
            const data = fs.readFileSync(outputPath);
            fs.unlinkSync(outputPath);
            resolve(data);
          } catch (err) {
            reject(err);
          }
        },
      );
    });
  }

  async generateAnimatedImage(
    videoPath: string,
    startTime: number,
    endTime: number,
    padding: number = 0.5,
    options: {
      fps?: number;
      maxWidth?: number;
      maxHeight?: number;
      crf?: number;
    } = {},
  ): Promise<Buffer> {
    const start = Math.max(0, startTime - padding);
    const duration = endTime - startTime + 2 * padding;
    const { fps = 10, maxWidth = 640, maxHeight, crf = 35 } = options;

    const clampedFps = Math.max(1, Math.min(60, fps));
    const clampedCrf = Math.max(0, Math.min(63, crf));

    const vfParts: string[] = [];
    vfParts.push(`fps=${clampedFps}`);
    if (maxWidth && maxWidth > 0 && maxHeight && maxHeight > 0) {
      vfParts.push(
        `scale=w=${maxWidth}:h=${maxHeight}:force_original_aspect_ratio=decrease`,
      );
    } else if (maxWidth && maxWidth > 0) {
      vfParts.push(`scale=w=${maxWidth}:h=-2`);
    } else if (maxHeight && maxHeight > 0) {
      vfParts.push(`scale=w=-2:h=${maxHeight}`);
    }

    return new Promise((resolve, reject) => {
      const outputPath = path.join(
        this.tempDir,
        `animation_${Date.now()}.avif`,
      );

      execFile(
        "ffmpeg",
        [
          "-ss",
          start.toString(),
          "-t",
          duration.toString(),
          "-i",
          videoPath,
          "-vf",
          vfParts.join(","),
          "-c:v",
          "libaom-av1",
          "-crf",
          clampedCrf.toString(),
          "-b:v",
          "0",
          "-cpu-used",
          "8",
          "-y",
          outputPath,
        ],
        { timeout: 60000 },
        (error) => {
          if (error) {
            reject(this.ffmpegError("animation generation", error));
            return;
          }

          try {
            const data = fs.readFileSync(outputPath);
            fs.unlinkSync(outputPath);
            resolve(data);
          } catch (err) {
            reject(err);
          }
        },
      );
    });
  }

  cleanup(): void {
    try {
      if (fs.existsSync(this.tempDir)) {
        fs.rmSync(this.tempDir, { recursive: true, force: true });
      }
    } catch (err) {
      console.error("Failed to cleanup media generator temp directory:", err);
    }
  }
}
