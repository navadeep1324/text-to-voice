import { NextRequest } from 'next/server';
import fs from 'fs';
import path from 'path';
import { randomBytes } from 'crypto';
import { loadConfig } from '@/lib/config';
import { TTSEngine } from '@/lib/ttsEngine';
import { jobStore } from '@/lib/jobStore';
import { parseSheetRows, buildSectionedVoiceText } from '@/lib/excelParser';

export const dynamic = 'force-dynamic';

async function processJob(
  jobId: string,
  filePath: string,
  sheetName: string,
  proto: string,
  host: string
) {
  try {
    if (!fs.existsSync(filePath)) {
      jobStore.set(jobId, { status: 'error', error: `File not found: ${filePath}` });
      return;
    }

    const fileBuffer = fs.readFileSync(filePath);
    const rows = parseSheetRows(fileBuffer, sheetName);

    if (rows.length < 2) {
      jobStore.set(jobId, { status: 'error', error: `Sheet '${sheetName}' has no data rows` });
      return;
    }

    const voiceText = buildSectionedVoiceText(rows);
    if (!voiceText) {
      jobStore.set(jobId, { status: 'error', error: `No text content found in sheet '${sheetName}'` });
      return;
    }

    const config = loadConfig();
    const tts = new TTSEngine(config.tts);

    const baseName = path.basename(filePath, path.extname(filePath));
    const fileId = `${baseName}__${sheetName}`;
    const audioPath = await tts.convert(voiceText, fileId);

    const audioBase64 = fs.readFileSync(audioPath).toString('base64');
    const filename = audioPath.split(/[\\/]/).pop();
    const audioUrl = `${proto}://${host}/audio/${filename}`;

    jobStore.set(jobId, {
      status: 'done',
      tab_name: sheetName,
      date_id: baseName,
      filename,
      voice_text: voiceText,
      audio_url: audioUrl,
      audio_base64: audioBase64,
    });

    console.log(`Job ${jobId} done: ${filename}`);
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    console.error(`Job ${jobId} failed:`, msg);
    jobStore.set(jobId, { status: 'error', error: msg });
  }
}

export async function POST(request: NextRequest) {
  let body: { file_path?: string; sheet_name?: string } = {};
  try {
    body = await request.json();
  } catch {
    return Response.json({ error: 'Request body must be valid JSON' }, { status: 400 });
  }

  const { file_path, sheet_name } = body;
  if (!file_path) return Response.json({ error: 'file_path is required' }, { status: 400 });
  if (!sheet_name) return Response.json({ error: 'sheet_name is required' }, { status: 400 });

  const jobId = randomBytes(8).toString('hex');
  jobStore.set(jobId, { status: 'processing' });

  const proto = request.headers.get('x-forwarded-proto') || 'https';
  const host = request.headers.get('host') || '';

  processJob(jobId, file_path, sheet_name, proto, host).catch(() => {});

  return Response.json({ job_id: jobId, status: 'processing' });
}
