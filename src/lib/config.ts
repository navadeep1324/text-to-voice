import fs from 'fs';
import path from 'path';

function loadLocalEnv() {
  const envPath = path.join(process.cwd(), '.env.local');
  if (!fs.existsSync(envPath)) return;

  for (const line of fs.readFileSync(envPath, 'utf-8').split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;

    const separatorIndex = trimmed.indexOf('=');
    if (separatorIndex === -1) continue;

    const key = trimmed.slice(0, separatorIndex).trim().replace(/^\uFEFF/, '');
    const value = trimmed.slice(separatorIndex + 1).trim();
    process.env[key] ??= value.replace(/^["']|["']$/g, '');
  }
}

export function loadConfig() {
  loadLocalEnv();

  // Try config.json first (local dev)
  const configPath = path.join(process.cwd(), 'config.json');
  if (fs.existsSync(configPath)) {
    const config = JSON.parse(fs.readFileSync(configPath, 'utf-8'));
    const spClientSecret = process.env.SP_CLIENT_SECRET;

    if (config.sharepoint) {
      config.sharepoint.client_secret = spClientSecret || config.sharepoint.client_secret;

      if (!config.sharepoint.client_secret) {
        throw new Error('SharePoint client secret is missing. Set SP_CLIENT_SECRET in the environment.');
      }
    }

    return config;
  }

  // Fall back to environment variables (production on Hostinger)
  const sheetId = process.env.SHEET_ID;
  const tabs = process.env.SHEET_TABS;
  const voice = process.env.TTS_VOICE;

  if (!sheetId || !tabs) {
    throw new Error('No config.json found and SHEET_ID / SHEET_TABS env vars are not set');
  }

  const spTenantId = process.env.SP_TENANT_ID;
  const spClientId = process.env.SP_CLIENT_ID;
  const spClientSecret = process.env.SP_CLIENT_SECRET;
  const spHost = process.env.SP_HOST;
  const spSitePath = process.env.SP_SITE_PATH;
  const spFilePath = process.env.SP_FILE_PATH;
  const spSheets = process.env.SP_SHEETS;

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
    schedule: {
      cron: process.env.SCHEDULE_CRON || '0 9 * * 1-5',
      timezone: process.env.SCHEDULE_TZ || 'Asia/Kolkata',
    },
    ...(spTenantId && spClientId && spClientSecret && spHost && spSitePath && spFilePath
      ? {
          sharepoint: {
            tenant_id: spTenantId,
            client_id: spClientId,
            client_secret: spClientSecret,
            host: spHost,
            site_path: spSitePath,
            file_path: spFilePath,
            sheets: spSheets ? spSheets.split(',').map((s) => s.trim()) : ['AI Team', 'Dev Team', 'DM Team'],
          },
        }
      : {}),
  };
}
