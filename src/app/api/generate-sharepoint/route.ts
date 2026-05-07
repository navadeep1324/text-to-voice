import { NextRequest } from 'next/server';
import fs from 'fs';
import { randomBytes } from 'crypto';
import { loadConfig } from '@/lib/config';
import { TTSEngine } from '@/lib/ttsEngine';
import { jobStore } from '@/lib/jobStore';
import { downloadSharePointFile } from '@/lib/sharePointReader';
import { parseSheetRows, buildSectionedVoiceText } from '@/lib/excelParser';

export const dynamic = 'force-dynamic';

export async function generateSharePointAudio(proto = 'http', host = 'localhost:8000') {
  const config = loadConfig();

  if (!config.sharepoint) {
    throw new Error('SharePoint is not configured. Add a "sharepoint" block to config.json.');
  }

  const sheets: string[] = config.sharepoint.sheets ?? ['AI Team', 'Dev Team', 'DM Team'];
  const results: { sheet: string; filename: string; audio_url: string }[] = [];

  // Download the file once, reuse for all sheets
  const fileBuffer = await downloadSharePointFile(config.sharepoint);
  const tts = new TTSEngine(config.tts);

  for (const sheetName of sheets) {
    try {
      const rows = parseSheetRows(fileBuffer, sheetName);
      if (rows.length < 2) {
        console.warn(`Sheet '${sheetName}' has no data rows — skipping`);
        continue;
      }

      const voiceText = buildSectionedVoiceText(rows);
      if (!voiceText) {
        console.warn(`Sheet '${sheetName}' produced no voice text — skipping`);
        continue;
      }

      const safeSheet = sheetName.replace(/\s+/g, '_');
      const fileId = `sharepoint__${safeSheet}`;
      const audioPath = await tts.convert(voiceText, fileId);
      const filename = audioPath.split(/[\\/]/).pop()!;
      const audioUrl = `${proto}://${host}/audio/${filename}`;

      results.push({ sheet: sheetName, filename, audio_url: audioUrl });
      console.log(`[SharePoint] Generated: ${filename}`);
    } catch (e) {
      console.error(`[SharePoint] Sheet '${sheetName}' failed:`, e instanceof Error ? e.message : e);
    }
  }

  return results;
}

async function processJob(jobId: string, sheetName: string | null, proto: string, host: string) {
  try {
    const config = loadConfig();
    if (!config.sharepoint) {
      jobStore.set(jobId, { status: 'error', error: 'SharePoint is not configured.' });
      return;
    }

    const sheets = sheetName
      ? [sheetName]
      : (config.sharepoint.sheets ?? ['AI Team', 'Dev Team', 'DM Team']);

    const fileBuffer = await downloadSharePointFile(config.sharepoint);
    const tts = new TTSEngine(config.tts);
    const outputs: { sheet: string; filename: string; audio_url: string }[] = [];

    for (const sheet of sheets) {
      const rows = parseSheetRows(fileBuffer, sheet);
      if (rows.length < 2) continue;
      const voiceText = buildSectionedVoiceText(rows);
      if (!voiceText) continue;

      const safeSheet = sheet.replace(/\s+/g, '_');
      const audioPath = await tts.convert(voiceText, `sharepoint__${safeSheet}`);
      const audioBase64 = fs.readFileSync(audioPath).toString('base64');
      const filename = audioPath.split(/[\\/]/).pop()!;
      outputs.push({ sheet, filename, audio_url: `${proto}://${host}/audio/${filename}` });

      // Store individual job entry for single-sheet requests
      if (sheets.length === 1) {
        jobStore.set(jobId, {
          status: 'done',
          tab_name: sheet,
          filename,
          voice_text: voiceText,
          audio_url: `${proto}://${host}/audio/${filename}`,
          audio_base64: audioBase64,
        });
      }
    }

    if (sheets.length > 1) {
      jobStore.set(jobId, { status: 'done', results: outputs });
    }

    console.log(`Job ${jobId} done: ${outputs.map((o) => o.filename).join(', ')}`);
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    console.error(`Job ${jobId} failed:`, msg);
    jobStore.set(jobId, { status: 'error', error: msg });
  }
}

export async function POST(request: NextRequest) {
  let body: { sheet_name?: string } = {};
  try { body = await request.json(); } catch { /* no body = all sheets */ }

  const jobId = randomBytes(8).toString('hex');
  jobStore.set(jobId, { status: 'processing' });

  const proto = request.headers.get('x-forwarded-proto') || 'http';
  const host = request.headers.get('host') || 'localhost:8000';

  processJob(jobId, body.sheet_name ?? null, proto, host).catch(() => {});

  return Response.json({ job_id: jobId, status: 'processing' });
}
