import WebSocket from 'ws';
import { createHash, randomBytes } from 'crypto';
import fs from 'fs';
import path from 'path';

// ── Constants (from Python edge_tts/constants.py) ────────────────────────────
const TRUSTED_CLIENT_TOKEN = '6A5AA1D4EAFF4E9FB37E23D68491D6F4';
const WSS_URL = `wss://speech.platform.bing.com/consumer/speech/synthesize/readaloud/edge/v1?TrustedClientToken=${TRUSTED_CLIENT_TOKEN}`;
const CHROMIUM_FULL_VERSION = '143.0.3650.75';
const CHROMIUM_MAJOR = '143';
const SEC_MS_GEC_VERSION = `1-${CHROMIUM_FULL_VERSION}`;

// ── DRM token (from Python edge_tts/drm.py) ──────────────────────────────────
function generateSecMsGec(): string {
  const WIN_EPOCH = 11644473600;
  const S_TO_100NS = 10_000_000;
  let ticks = Date.now() / 1000 + WIN_EPOCH;
  ticks -= ticks % 300;
  ticks = Math.round(ticks * S_TO_100NS);
  return createHash('sha256').update(`${ticks}${TRUSTED_CLIENT_TOKEN}`, 'ascii').digest('hex').toUpperCase();
}

function generateMuid(): string {
  return randomBytes(16).toString('hex').toUpperCase();
}

// ── Headers (from Python edge_tts/constants.py) ───────────────────────────────
function buildWssHeaders(): Record<string, string> {
  return {
    'Pragma': 'no-cache',
    'Cache-Control': 'no-cache',
    'Origin': 'chrome-extension://jdiccldimpdaibmpdkjnbmckianbfold',
    'Sec-WebSocket-Version': '13',
    'User-Agent': `Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/${CHROMIUM_MAJOR}.0.0.0 Safari/537.36 Edg/${CHROMIUM_MAJOR}.0.0.0`,
    'Accept-Encoding': 'gzip, deflate, br, zstd',
    'Accept-Language': 'en-US,en;q=0.9',
    'Cookie': `muid=${generateMuid()};`,
  };
}

// ── Date string (from Python edge_tts/communicate.py date_to_string) ─────────
function dateToString(): string {
  const now = new Date();
  const DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const p = (n: number) => String(n).padStart(2, '0');
  return `${DAYS[now.getUTCDay()]} ${MONTHS[now.getUTCMonth()]} ${p(now.getUTCDate())} ${now.getUTCFullYear()} ${p(now.getUTCHours())}:${p(now.getUTCMinutes())}:${p(now.getUTCSeconds())} GMT+0000 (Coordinated Universal Time)`;
}

function xmlEscape(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ── Speech config (from Python communicate.py send_command_request) ───────────
// Default boundary=SentenceBoundary → sentenceBoundaryEnabled="true", wordBoundaryEnabled="false"
function speechConfigMsg(): string {
  return (
    `X-Timestamp:${dateToString()}\r\n` +
    `Content-Type:application/json; charset=utf-8\r\n` +
    `Path:speech.config\r\n\r\n` +
    `{"context":{"synthesis":{"audio":{"metadataoptions":{"sentenceBoundaryEnabled":"true","wordBoundaryEnabled":"false"},"outputFormat":"audio-24khz-48kbitrate-mono-mp3"}}}}\r\n`
  );
}

// ── SSML message (from Python communicate.py ssml_headers_plus_data + mkssml) ─
function ssmlMsg(reqId: string, text: string, voice: string, rate: string, volume: string, pitch: string): string {
  const ssml =
    `<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='en-US'>` +
    `<voice name='${voice}'>` +
    `<prosody pitch='${pitch}' rate='${rate}' volume='${volume}'>${xmlEscape(text)}</prosody>` +
    `</voice></speak>`;
  return (
    `X-RequestId:${reqId}\r\n` +
    `Content-Type:application/ssml+xml\r\n` +
    `X-Timestamp:${dateToString()}Z\r\n` +
    `Path:ssml\r\n\r\n${ssml}`
  );
}

// ── Core synthesizer ──────────────────────────────────────────────────────────
async function synthesize(
  text: string,
  voice: string,
  rate: string,
  volume: string,
  pitch = '+0Hz'
): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const connectionId = randomBytes(16).toString('hex');
    const url = `${WSS_URL}&ConnectionId=${connectionId}&Sec-MS-GEC=${generateSecMsGec()}&Sec-MS-GEC-Version=${SEC_MS_GEC_VERSION}`;

    const ws = new WebSocket(url, {
      headers: buildWssHeaders(),
      perMessageDeflate: true,
    });

    const chunks: Buffer[] = [];
    let settled = false;

    const done = (buf: Buffer) => { if (!settled) { settled = true; clearTimeout(timer); ws.terminate(); resolve(buf); } };
    const fail = (err: Error)  => { if (!settled) { settled = true; clearTimeout(timer); ws.terminate(); reject(err); } };

    const timer = setTimeout(() => fail(new Error('TTS timed out after 60s — check if outbound WebSocket is allowed on this host')), 60_000);

    ws.on('open', () => {
      const reqId = randomBytes(16).toString('hex');
      ws.send(speechConfigMsg());
      ws.send(ssmlMsg(reqId, text, voice, rate, volume, pitch));
    });

    ws.on('message', (rawData: WebSocket.RawData, isBinary: boolean) => {
      if (isBinary) {
        try {
          const data = Buffer.isBuffer(rawData) ? rawData : Buffer.from(rawData as ArrayBuffer);
          const headerLen = data.readUInt16BE(0);
          const header = data.slice(2, headerLen + 2).toString('utf-8');
          const audio = data.slice(headerLen + 2);
          if (header.includes('Path:audio') && audio.length > 0) {
            chunks.push(Buffer.from(audio));
          }
        } catch { /* skip malformed frame */ }
      } else {
        const msg = Buffer.isBuffer(rawData) ? rawData.toString() : String(rawData);
        if (msg.includes('Path:turn.end')) {
          done(Buffer.concat(chunks));
        }
      }
    });

    ws.on('error', (err) => fail(new Error(`TTS WebSocket error: ${err.message}`)));
    ws.on('close', (code) => {
      if (chunks.length > 0) {
        done(Buffer.concat(chunks)); // resolve with what we have
      } else {
        fail(new Error(`TTS closed with no audio (code ${code})`));
      }
    });
  });
}

// ── TTSEngine class ───────────────────────────────────────────────────────────
interface TTSConfig {
  voice: string;
  rate: string;
  volume: string;
  output_folder: string;
}

export class TTSEngine {
  private voice: string;
  private rate: string;
  private volume: string;
  private outputFolder: string;

  constructor(config: TTSConfig) {
    this.voice = config.voice || 'en-US-JennyNeural';
    this.rate = config.rate || '+0%';
    this.volume = config.volume || '+0%';
    this.outputFolder = path.join(process.cwd(), config.output_folder || 'output/audio');
    fs.mkdirSync(this.outputFolder, { recursive: true });
  }

  async convert(text: string, rowId: string): Promise<string> {
    const safeId = rowId.replace(/[^\w\-]/g, '_');
    const outputPath = path.join(this.outputFolder, `${safeId}.mp3`);

    const audioBuffer = await synthesize(text, this.voice, this.rate, this.volume);
    fs.writeFileSync(outputPath, audioBuffer);

    console.log(`Audio saved: ${path.basename(outputPath)} (${audioBuffer.length} bytes)`);
    return outputPath;
  }
}
