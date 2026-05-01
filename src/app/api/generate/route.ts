import { NextRequest } from 'next/server';
import fs from 'fs';
import { randomBytes } from 'crypto';
import { loadConfig } from '@/lib/config';
import { GoogleSheetReader } from '@/lib/googleSheetReader';
import { TTSEngine } from '@/lib/ttsEngine';
import { jobStore } from '@/lib/jobStore';

export const dynamic = 'force-dynamic';

async function processJob(
  jobId: string,
  tabName: string,
  sheetId: string | undefined,
  dateFilter: string | undefined,
  proto: string,
  host: string
) {
  try {
    const config = loadConfig();
    const tts = new TTSEngine(config.tts);

    const gsConfig = { ...config.google_sheets, tabs: [tabName] };
    if (sheetId) gsConfig.sheet_id = sheetId;

    const reader = new GoogleSheetReader(gsConfig);
    let items = await reader.readAllTabs();

    if (!items.length) {
      jobStore.set(jobId, { status: 'error', error: `No data found in tab '${tabName}'` });
      return;
    }

    if (dateFilter) {
      items = items.filter(([, dateId]) => dateId.includes(dateFilter.replace(/-/g, '_')));
      if (!items.length) {
        jobStore.set(jobId, { status: 'error', error: `No data for date '${dateFilter}' in tab '${tabName}'` });
        return;
      }
    }

    const [tab, dateId, voiceText] = items[items.length - 1];
    const fileId = `${tab}__${dateId}`;

    const audioPath = await tts.convert(voiceText, fileId);
    const audioBase64 = fs.readFileSync(audioPath).toString('base64');
    const filename = audioPath.split(/[\\/]/).pop();
    const audioUrl = `${proto}://${host}/audio/${filename}`;

    jobStore.set(jobId, {
      status: 'done',
      tab_name: tab,
      date_id: dateId,
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
  let body: { tab_name?: string; sheet_id?: string; date_filter?: string } = {};
  try {
    body = await request.json();
  } catch {
    return Response.json({ error: 'Request body must be valid JSON' }, { status: 400 });
  }

  const { tab_name, sheet_id, date_filter } = body;
  if (!tab_name) {
    return Response.json({ error: 'tab_name is required' }, { status: 400 });
  }

  const jobId = randomBytes(8).toString('hex');
  jobStore.set(jobId, { status: 'processing' });

  const proto = request.headers.get('x-forwarded-proto') || 'https';
  const host = request.headers.get('host') || '';

  // Fire and forget — responds immediately, TTS runs in background
  processJob(jobId, tab_name, sheet_id, date_filter, proto, host).catch(() => {});

  return Response.json({ job_id: jobId, status: 'processing' });
}
