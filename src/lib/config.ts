import fs from 'fs';
import path from 'path';

export function loadConfig() {
  // Try config.json first (local dev)
  const configPath = path.join(process.cwd(), 'config.json');
  if (fs.existsSync(configPath)) {
    return JSON.parse(fs.readFileSync(configPath, 'utf-8'));
  }

  // Fall back to environment variables (production on Hostinger)
  const sheetId = process.env.SHEET_ID;
  const tabs = process.env.SHEET_TABS;
  const voice = process.env.TTS_VOICE;

  if (!sheetId || !tabs) {
    throw new Error('No config.json found and SHEET_ID / SHEET_TABS env vars are not set');
  }

  return {
    google_sheets: {
      sheet_id: sheetId,
      tabs: tabs.split(',').map((t) => t.trim()),
      date_column: process.env.DATE_COLUMN || 'Date',
      tasks_column: process.env.TASKS_COLUMN || 'Tasks',
      targeted_date_column: process.env.TARGETED_DATE_COLUMN || 'Targeted Date',
      remarks_column: process.env.REMARKS_COLUMN || 'Remarks',
    },
    tts: {
      voice: voice || 'en-IN-NeerjaExpressiveNeural',
      rate: process.env.TTS_RATE || '+0%',
      volume: process.env.TTS_VOLUME || '+0%',
      output_folder: process.env.OUTPUT_FOLDER || 'public/audio',
    },
  };
}
