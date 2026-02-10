import * as fs from "fs";
import * as http from "http";
import * as path from "path";

export class TexthookerService {
  private server: http.Server | null = null;

  public isRunning(): boolean {
    return this.server !== null;
  }

  public start(port: number): http.Server | null {
    const texthookerPath = this.getTexthookerPath();
    if (!texthookerPath) {
      console.error("texthooker-ui not found");
      return null;
    }

    this.server = http.createServer((req, res) => {
      const urlPath = (req.url || "/").split("?")[0];
      const filePath = path.join(
        texthookerPath,
        urlPath === "/" ? "index.html" : urlPath,
      );

      const ext = path.extname(filePath);
      const mimeTypes: Record<string, string> = {
        ".html": "text/html",
        ".js": "application/javascript",
        ".css": "text/css",
        ".json": "application/json",
        ".png": "image/png",
        ".svg": "image/svg+xml",
        ".ttf": "font/ttf",
        ".woff": "font/woff",
        ".woff2": "font/woff2",
      };

      fs.readFile(filePath, (err, data) => {
        if (err) {
          res.writeHead(404);
          res.end("Not found");
          return;
        }
        res.writeHead(200, { "Content-Type": mimeTypes[ext] || "text/plain" });
        res.end(data);
      });
    });

    this.server.listen(port, "127.0.0.1", () => {
      console.log(`Texthooker server running at http://127.0.0.1:${port}`);
    });

    return this.server;
  }

  public stop(): void {
    if (this.server) {
      this.server.close();
      this.server = null;
    }
  }

  private getTexthookerPath(): string | null {
    const searchPaths = [
      path.join(__dirname, "..", "..", "..", "vendor", "texthooker-ui", "docs"),
      path.join(
        process.resourcesPath,
        "app",
        "vendor",
        "texthooker-ui",
        "docs",
      ),
    ];
    for (const candidate of searchPaths) {
      if (fs.existsSync(path.join(candidate, "index.html"))) {
        return candidate;
      }
    }
    return null;
  }
}
