import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import WebSocket from "ws";

export function hasMpvWebsocketPlugin(): boolean {
  const mpvWebsocketPath = path.join(
    os.homedir(),
    ".config",
    "mpv",
    "mpv_websocket",
  );
  return fs.existsSync(mpvWebsocketPath);
}

export class SubtitleWebSocketService {
  private server: WebSocket.Server | null = null;

  public isRunning(): boolean {
    return this.server !== null;
  }

  public start(port: number, getCurrentSubtitleText: () => string): void {
    this.server = new WebSocket.Server({ port, host: "127.0.0.1" });

    this.server.on("connection", (ws: WebSocket) => {
      console.log("WebSocket client connected");
      const currentText = getCurrentSubtitleText();
      if (currentText) {
        ws.send(JSON.stringify({ sentence: currentText }));
      }
    });

    this.server.on("error", (err: Error) => {
      console.error("WebSocket server error:", err.message);
    });

    console.log(`Subtitle WebSocket server running on ws://127.0.0.1:${port}`);
  }

  public broadcast(text: string): void {
    if (!this.server) return;
    const message = JSON.stringify({ sentence: text });
    for (const client of this.server.clients) {
      if (client.readyState === WebSocket.OPEN) {
        client.send(message);
      }
    }
  }

  public stop(): void {
    if (this.server) {
      this.server.close();
      this.server = null;
    }
  }
}
