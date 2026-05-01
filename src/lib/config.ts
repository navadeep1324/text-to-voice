import fs from 'fs';
import path from 'path';

export function loadConfig() {
  const configPath = path.join(process.cwd(), 'config.json');
  const raw = fs.readFileSync(configPath, 'utf-8');
  return JSON.parse(raw);
}
