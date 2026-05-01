import fs from 'fs';
import { loadConfig } from '@/lib/config';
import { GoogleSheetReader } from '@/lib/googleSheetReader';
import { TTSEngine } from '@/lib/ttsEngine';

const config = loadConfig();
const tts = new TTSEngine(config.tts);

export async function POST() {
  const reader = new GoogleSheetReader(config.google_sheets);
  let items;
  try {
    items = await reader.readAllTabs();
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    return Response.json({ error: `Sheet read failed: ${msg}` }, { status: 500 });
  }

  const results = [];
  for (const [tabName, dateId, voiceText] of items) {
    const fileId = `${tabName}__${dateId}`;
    try {
      const audioPath = await tts.convert(voiceText, fileId);
      const audioBase64 = fs.readFileSync(audioPath).toString('base64');
      const filename = audioPath.split(/[\\/]/).pop();
      results.push({ tab_name: tabName, date_id: dateId, filename, voice_text: voiceText, audio_base64: audioBase64 });
      console.log(`Generated: ${filename}`);
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      console.error(`Failed ${fileId}:`, msg);
      results.push({ tab_name: tabName, date_id: dateId, error: msg });
    }
  }

  return Response.json({ count: results.length, results });
}
