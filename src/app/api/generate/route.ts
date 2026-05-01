import { NextRequest } from 'next/server';
import fs from 'fs';
import { loadConfig } from '@/lib/config';
import { GoogleSheetReader } from '@/lib/googleSheetReader';
import { TTSEngine } from '@/lib/ttsEngine';

export const dynamic = 'force-dynamic';

export async function POST(request: NextRequest) {
  const config = loadConfig();
  const tts = new TTSEngine(config.tts);

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

  const gsConfig = { ...config.google_sheets, tabs: [tab_name] };
  if (sheet_id) gsConfig.sheet_id = sheet_id;

  const reader = new GoogleSheetReader(gsConfig);
  let items;
  try {
    items = await reader.readAllTabs();
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    return Response.json({ error: `Sheet read failed: ${msg}` }, { status: 500 });
  }

  if (!items.length) {
    return Response.json({ error: `No data found in tab '${tab_name}'` }, { status: 404 });
  }

  if (date_filter) {
    items = items.filter(([, dateId]) => dateId.includes(date_filter.replace(/-/g, '_')));
    if (!items.length) {
      return Response.json(
        { error: `No data for date '${date_filter}' in tab '${tab_name}'` },
        { status: 404 }
      );
    }
  }

  const [tabName, dateId, voiceText] = items[items.length - 1];
  const fileId = `${tabName}__${dateId}`;

  let audioPath: string;
  try {
    audioPath = await tts.convert(voiceText, fileId);
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    console.error('TTS error:', e);
    return Response.json({ error: `TTS failed: ${msg}` }, { status: 500 });
  }

  const audioBase64 = fs.readFileSync(audioPath).toString('base64');
  const filename = audioPath.split(/[\\/]/).pop();

  return Response.json({ tab_name: tabName, date_id: dateId, filename, voice_text: voiceText, audio_base64: audioBase64 });
}
