import * as fs from "fs";
import * as path from "path";
import { DEFAULT_CONFIG, generateConfigTemplate } from "./config";

function main(): void {
  const outputPath = path.join(process.cwd(), "config.example.jsonc");
  const template = generateConfigTemplate(DEFAULT_CONFIG);
  fs.writeFileSync(outputPath, template, "utf-8");
  console.log(`Generated ${outputPath}`);
}

main();
