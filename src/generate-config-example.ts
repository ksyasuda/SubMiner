import * as fs from 'fs';
import * as path from 'path';
import { DEFAULT_CONFIG, generateConfigTemplate } from './config';

function main(): void {
  const template = generateConfigTemplate(DEFAULT_CONFIG);
  const outputPaths = [
    path.join(process.cwd(), 'config.example.jsonc'),
    path.join(process.cwd(), 'docs', 'public', 'config.example.jsonc'),
  ];

  for (const outputPath of outputPaths) {
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    fs.writeFileSync(outputPath, template, 'utf-8');
    console.log(`Generated ${outputPath}`);
  }
}

main();
