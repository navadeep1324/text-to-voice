const fs = require('fs');
const path = require('path');
const { createHash, randomBytes } = require('crypto');
const XLSX = require('xlsx');
const WebSocket = require('ws');

const TRUSTED_CLIENT_TOKEN = '6A5AA1D4EAFF4E9FB37E23D68491D6F4';
const CHROMIUM_FULL_VERSION = '143.0.3650.75';
const CHROMIUM_MAJOR = '143';
const SEC_MS_GEC_VERSION = `1-${CHROMIUM_FULL_VERSION}`;

function loadLocalEnv() {
  const envPath = path.join(process.cwd(), '.env.local');
  if (!fs.existsSync(envPath)) return {};

  const env = {};
  for (const line of fs.readFileSync(envPath, 'utf-8').split(/\r?\n/)) {
    const trimmed = line.trim().replace(/^\uFEFF/, '');
    if (!trimmed || trimmed.startsWith('#')) continue;
    const separatorIndex = trimmed.indexOf('=');
    if (separatorIndex === -1) continue;
    env[trimmed.slice(0, separatorIndex).trim()] = trimmed.slice(separatorIndex + 1).trim();
  }
  return env;
}

function splitCsv(value, fallback = []) {
  if (!value) return fallback;
  return value.split(',').map((item) => item.trim()).filter(Boolean);
}

function loadConfig(env) {
  const configPath = path.join(process.cwd(), 'config.json');
  if (fs.existsSync(configPath)) {
    return JSON.parse(fs.readFileSync(configPath, 'utf-8'));
  }

  const required = [
    'SP_TENANT_ID',
    'SP_CLIENT_ID',
    'SP_HOST',
    'SP_SITE_PATH',
    'SP_FILE_PATH',
  ];
  const missing = required.filter((key) => !process.env[key] && !env[key]);
  if (missing.length > 0) {
    throw new Error(`config.json not found and missing environment variables: ${missing.join(', ')}`);
  }

  return {
    sharepoint: {
      tenant_id: process.env.SP_TENANT_ID || env.SP_TENANT_ID,
      client_id: process.env.SP_CLIENT_ID || env.SP_CLIENT_ID,
      host: process.env.SP_HOST || env.SP_HOST,
      site_path: process.env.SP_SITE_PATH || env.SP_SITE_PATH,
      file_path: process.env.SP_FILE_PATH || env.SP_FILE_PATH,
      sheets: splitCsv(process.env.SP_SHEETS || env.SP_SHEETS, ['AI Team', 'Dev Team', 'DM Team']),
    },
    tts: {
      voice: process.env.TTS_VOICE || env.TTS_VOICE || 'en-IN-NeerjaExpressiveNeural',
      rate: process.env.TTS_RATE || env.TTS_RATE || '+0%',
      volume: process.env.TTS_VOLUME || env.TTS_VOLUME || '+0%',
      output_folder: process.env.OUTPUT_FOLDER || env.OUTPUT_FOLDER || 'public/audio',
    },
    teams: {
      public_base_url: process.env.PUBLIC_BASE_URL || env.PUBLIC_BASE_URL,
    },
  };
}

function dateToString() {
  const now = new Date();
  const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const pad = (n) => String(n).padStart(2, '0');
  return `${days[now.getUTCDay()]} ${months[now.getUTCMonth()]} ${pad(now.getUTCDate())} ${now.getUTCFullYear()} ${pad(now.getUTCHours())}:${pad(now.getUTCMinutes())}:${pad(now.getUTCSeconds())} GMT+0000 (Coordinated Universal Time)`;
}

function generateSecMsGec() {
  const winEpoch = 11644473600;
  const sTo100ns = 10_000_000;
  let ticks = Date.now() / 1000 + winEpoch;
  ticks -= ticks % 300;
  ticks = Math.round(ticks * sTo100ns);
  return createHash('sha256').update(`${ticks}${TRUSTED_CLIENT_TOKEN}`, 'ascii').digest('hex').toUpperCase();
}

function xmlEscape(text) {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function buildWssHeaders() {
  return {
    Pragma: 'no-cache',
    'Cache-Control': 'no-cache',
    Origin: 'chrome-extension://jdiccldimpdaibmpdkjnbmckianbfold',
    'Sec-WebSocket-Version': '13',
    'User-Agent': `Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/${CHROMIUM_MAJOR}.0.0.0 Safari/537.36 Edg/${CHROMIUM_MAJOR}.0.0.0`,
    'Accept-Encoding': 'gzip, deflate, br, zstd',
    'Accept-Language': 'en-US,en;q=0.9',
    Cookie: `muid=${randomBytes(16).toString('hex').toUpperCase()};`,
  };
}

function speechConfigMsg() {
  return [
    `X-Timestamp:${dateToString()}`,
    'Content-Type:application/json; charset=utf-8',
    'Path:speech.config',
    '',
    JSON.stringify({
      context: {
        synthesis: {
          audio: {
            metadataoptions: {
              sentenceBoundaryEnabled: 'true',
              wordBoundaryEnabled: 'false',
            },
            outputFormat: 'audio-24khz-48kbitrate-mono-mp3',
          },
        },
      },
    }),
    '',
  ].join('\r\n');
}

function ssmlMsg(reqId, text, voice, rate, volume, pitch = '+0Hz') {
  const ssml = `<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='en-US'><voice name='${voice}'><prosody pitch='${pitch}' rate='${rate}' volume='${volume}'>${xmlEscape(text)}</prosody></voice></speak>`;
  return [
    `X-RequestId:${reqId}`,
    'Content-Type:application/ssml+xml',
    `X-Timestamp:${dateToString()}Z`,
    'Path:ssml',
    '',
    ssml,
  ].join('\r\n');
}

function synthesize(text, voice, rate, volume) {
  return new Promise((resolve, reject) => {
    const connectionId = randomBytes(16).toString('hex');
    const url = `wss://speech.platform.bing.com/consumer/speech/synthesize/readaloud/edge/v1?TrustedClientToken=${TRUSTED_CLIENT_TOKEN}&ConnectionId=${connectionId}&Sec-MS-GEC=${generateSecMsGec()}&Sec-MS-GEC-Version=${SEC_MS_GEC_VERSION}`;
    const ws = new WebSocket(url, { headers: buildWssHeaders(), perMessageDeflate: true });
    const chunks = [];
    let settled = false;

    const timer = setTimeout(() => fail(new Error('TTS timed out after 60s')), 60_000);
    const done = (buffer) => {
      if (!settled) {
        settled = true;
        clearTimeout(timer);
        ws.terminate();
        resolve(buffer);
      }
    };
    const fail = (error) => {
      if (!settled) {
        settled = true;
        clearTimeout(timer);
        try { ws.terminate(); } catch {}
        reject(error);
      }
    };

    ws.on('open', () => {
      const reqId = randomBytes(16).toString('hex');
      ws.send(speechConfigMsg());
      ws.send(ssmlMsg(reqId, text, voice, rate, volume));
    });

    ws.on('message', (rawData, isBinary) => {
      if (isBinary) {
        try {
          const data = Buffer.isBuffer(rawData) ? rawData : Buffer.from(rawData);
          const headerLen = data.readUInt16BE(0);
          const header = data.slice(2, headerLen + 2).toString('utf-8');
          const audio = data.slice(headerLen + 2);
          if (header.includes('Path:audio') && audio.length > 0) chunks.push(Buffer.from(audio));
        } catch {}
      } else {
        const message = Buffer.isBuffer(rawData) ? rawData.toString() : String(rawData);
        if (message.includes('Path:turn.end')) done(Buffer.concat(chunks));
      }
    });

    ws.on('error', (error) => fail(new Error(`TTS WebSocket error: ${error.message}`)));
    ws.on('close', (code) => {
      if (chunks.length > 0) done(Buffer.concat(chunks));
      else fail(new Error(`TTS closed with no audio (code ${code})`));
    });
  });
}

function buildSectionedVoiceText(rows, includeDateIntro = true) {
  const sections = [];
  let dateIntro = '';

  for (let i = 1; i < rows.length; i++) {
    const col0 = (rows[i][0] ?? '').toString().trim();
    const col1 = (rows[i][1] ?? '').toString().trim();
    const col2 = (rows[i][2] ?? '').toString().trim();
    const col3 = (rows[i][3] ?? '').toString().trim();

    if (col0 && !col1) {
      if (col0 === col0.toUpperCase() && col0.replace(/[^A-Z]/g, '').length > 2) {
        sections.push({ header: col0, sub: '', tasks: [] });
      } else if (sections.length === 0 && !dateIntro) {
        dateIntro = col0;
      } else if (sections.length > 0) {
        sections[sections.length - 1].sub = col0;
      }
      continue;
    }

    if (!col1) continue;

    let segment = col1.endsWith('.') ? col1 : `${col1}.`;
    if (col2) segment += ` Targeted date ${col2.replace(/\n/g, ' ')}.`;
    if (col3) segment += ` Status: ${col3}.`;

    if (sections.length === 0) sections.push({ header: '', sub: '', tasks: [] });
    sections[sections.length - 1].tasks.push(segment);
  }

  const parts = [];
  if (includeDateIntro && dateIntro) parts.push(`${dateIntro}.`);
  for (const section of sections) {
    if (section.tasks.length === 0) continue;
    if (section.header) parts.push(`${section.header}.`);
    if (section.sub) parts.push(`${section.sub}.`);
    parts.push(...section.tasks);
  }
  return parts.join(' ');
}

function dateLabelForIst(date = new Date()) {
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Kolkata',
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  }).formatToParts(date);
  const day = Number(parts.find((part) => part.type === 'day')?.value);
  const month = parts.find((part) => part.type === 'month')?.value;
  const year = parts.find((part) => part.type === 'year')?.value;
  const suffix = day % 10 === 1 && day % 100 !== 11
    ? 'st'
    : day % 10 === 2 && day % 100 !== 12
      ? 'nd'
      : day % 10 === 3 && day % 100 !== 13
        ? 'rd'
        : 'th';
  return `${day}${suffix} ${month} ${year}`;
}

function normalizeDateLabel(value) {
  if (value == null) return '';
  if (value instanceof Date && !Number.isNaN(value.getTime())) return dateLabelForIst(value);
  return String(value)
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/(\d+)(st|nd|rd|th)/i, '$1')
    .toLowerCase();
}

function hasCurrentDateUpdate(rows, todayLabel) {
  const rowAfterHeader = rows[1] || [];
  return normalizeDateLabel(rowAfterHeader[0]) === normalizeDateLabel(todayLabel);
}


function encodePath(filePath) {
  return filePath.split('/').map(encodeURIComponent).join('/');
}

function audioUrl(baseUrl, filename) {
  if (!baseUrl) return null;
  return `${baseUrl.replace(/\/$/, '')}/audio/${encodeURIComponent(filename)}`;
}

async function sendTeamsNotification(webhookUrl, config, results) {
  if (!webhookUrl) {
    console.log('TEAMS_WEBHOOK_URL is not set. Skipping Teams notification.');
    return;
  }

  const publicBaseUrl = config.teams?.public_base_url || process.env.PUBLIC_BASE_URL;
  const facts = results.map((result) => ({
    title: result.sheet,
    value: audioUrl(publicBaseUrl, result.filename) || result.filename,
  }));

  const payload = {
    type: 'message',
    attachments: [
      {
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: {
          $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
          type: 'AdaptiveCard',
          version: '1.4',
          body: [
            {
              type: 'TextBlock',
              text: 'Daily Excel Audio Generated',
              weight: 'Bolder',
              size: 'Medium',
            },
            {
              type: 'TextBlock',
              text: `Generated ${results.length} audio file(s) from SharePoint Excel.`,
              wrap: true,
            },
            {
              type: 'FactSet',
              facts,
            },
          ],
        },
      },
    ],
  };

  const response = await fetch(webhookUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    throw new Error(`Teams notification failed (${response.status}): ${await response.text()}`);
  }

  console.log('Teams notification sent.');
}

async function getSharePointFile(config) {
  const body = new URLSearchParams({
    client_id: config.client_id,
    client_secret: config.client_secret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });
  const tokenRes = await fetch(`https://login.microsoftonline.com/${config.tenant_id}/oauth2/v2.0/token`, { method: 'POST', body });
  const tokenJson = await tokenRes.json();
  if (!tokenRes.ok) throw new Error(`SharePoint auth failed (${tokenRes.status}): ${JSON.stringify(tokenJson)}`);

  const headers = { Authorization: `Bearer ${tokenJson.access_token}` };
  const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${config.host}:${config.site_path}`, { headers });
  const siteJson = await siteRes.json();
  if (!siteRes.ok) throw new Error(`SharePoint site failed (${siteRes.status}): ${JSON.stringify(siteJson)}`);

  const contentRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteJson.id}/drive/root:/${encodePath(config.file_path)}:/content`,
    { headers, redirect: 'follow' }
  );
  const buffer = Buffer.from(await contentRes.arrayBuffer());
  if (!contentRes.ok) throw new Error(`SharePoint file download failed (${contentRes.status}): ${buffer.toString('utf8', 0, 500)}`);
  return { buffer, siteId: siteJson.id };
}

async function uploadSharePointFile(config, siteId, buffer) {
  const body = new URLSearchParams({
    client_id: config.client_id,
    client_secret: config.client_secret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });
  const tokenRes = await fetch(`https://login.microsoftonline.com/${config.tenant_id}/oauth2/v2.0/token`, { method: 'POST', body });
  const tokenJson = await tokenRes.json();
  if (!tokenRes.ok) throw new Error(`SharePoint auth failed (${tokenRes.status}): ${JSON.stringify(tokenJson)}`);

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodePath(config.file_path)}:/content`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${tokenJson.access_token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      body: buffer,
    }
  );

  if (!response.ok) {
    throw new Error(`SharePoint file upload failed (${response.status}): ${await response.text()}`);
  }
}

async function runJob(options = {}) {
  const persistAudio = options.persistAudio ?? true;
  const notifyTeams = options.notifyTeams ?? true;
  const includeBase64 = options.includeBase64 ?? false;
  const env = loadLocalEnv();
  const config = loadConfig(env);
  const sharepoint = { ...config.sharepoint, client_secret: process.env.SP_CLIENT_SECRET || env.SP_CLIENT_SECRET };
  if (!sharepoint.client_secret) throw new Error('SP_CLIENT_SECRET is missing. Add it to .env.local.');

  const { buffer: fileBuffer, siteId } = await getSharePointFile(sharepoint);
  const workbook = XLSX.read(fileBuffer, { type: 'buffer', cellDates: true });
  const outputFolder = path.join(process.cwd(), config.tts?.output_folder || 'public/audio');
  fs.mkdirSync(outputFolder, { recursive: true });

  const results = [];
  const skipped = [];
  const todayLabel = dateLabelForIst();
  const voiceSegments = [];

  for (const sheetName of sharepoint.sheets || workbook.SheetNames) {
    if (!workbook.SheetNames.includes(sheetName)) {
      throw new Error(`Sheet '${sheetName}' not found. Available: ${workbook.SheetNames.join(', ')}`);
    }

    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: null, raw: false });
    if (!hasCurrentDateUpdate(rows, todayLabel)) {
      skipped.push({ sheet: sheetName, reason: `row 2 date is not ${todayLabel}` });
      console.log(`Skipping ${sheetName}: row 2 date is not ${todayLabel}`);
      continue;
    }

    const text = buildSectionedVoiceText(rows, voiceSegments.length === 0);
    if (!text) {
      skipped.push({ sheet: sheetName, reason: 'no voice text' });
      console.log(`Skipping ${sheetName}: no voice text`);
      continue;
    }

    voiceSegments.push({ sheetName, text: `Updates from ${sheetName}. ${text}` });
  }

  if (voiceSegments.length === 0) {
    console.log('No current-date updates found. Audio generation and Teams notification skipped.');
    return { result: 'OK', workbook: sharepoint.file_path, today: todayLabel, results: [], skipped };
  }

  const combinedText = voiceSegments.map(s => s.text).join(' ');
  const includedSheets = voiceSegments.map(s => s.sheetName).join(', ');
  const safeDate = todayLabel.replace(/\s+/g, '_').replace(/[^\w-]/g, '');
  const filename = `sharepoint__daily_update_${safeDate}.mp3`;

  console.log(`Synthesizing combined audio for [${includedSheets}] (${combinedText.length} chars)`);
  const audioBuffer = await synthesize(
    combinedText,
    config.tts?.voice || 'en-IN-NeerjaExpressiveNeural',
    config.tts?.rate || '+0%',
    config.tts?.volume || '+0%'
  );

  const result = { sheet: includedSheets, filename, bytes: audioBuffer.length };

  if (persistAudio) {
    const outputPath = path.join(outputFolder, filename);
    fs.writeFileSync(outputPath, audioBuffer);
    result.outputPath = outputPath;
    console.log(`Saved ${filename} (${audioBuffer.length} bytes)`);
  } else {
    console.log(`Generated ${filename} (${audioBuffer.length} bytes, not saved to disk)`);
  }

  if (includeBase64) {
    result.contentType = 'audio/mpeg';
    result.base64 = audioBuffer.toString('base64');
  }

  results.push(result);

  if (notifyTeams) {
    await sendTeamsNotification(process.env.TEAMS_WEBHOOK_URL || env.TEAMS_WEBHOOK_URL, config, results);
  }

  return { result: 'OK', workbook: sharepoint.file_path, today: todayLabel, results, skipped };
}

async function checkSheets() {
  const env = loadLocalEnv();
  const config = loadConfig(env);
  const sharepoint = { ...config.sharepoint, client_secret: process.env.SP_CLIENT_SECRET || env.SP_CLIENT_SECRET };
  if (!sharepoint.client_secret) throw new Error('SP_CLIENT_SECRET is missing. Add it to .env.local.');

  const { buffer: fileBuffer } = await getSharePointFile(sharepoint);
  const workbook = XLSX.read(fileBuffer, { type: 'buffer', cellDates: true });
  const todayLabel = dateLabelForIst();

  const filled = [];
  const unfilled = [];

  for (const sheetName of sharepoint.sheets || workbook.SheetNames) {
    if (!workbook.SheetNames.includes(sheetName)) continue;
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: null, raw: false });
    if (hasCurrentDateUpdate(rows, todayLabel)) {
      filled.push(sheetName);
    } else {
      unfilled.push(sheetName);
    }
  }

  return { today: todayLabel, filled, unfilled };
}

async function main() {
  const output = await runJob();
  console.log(JSON.stringify(output, null, 2));
}

if (require.main === module) {
  main().catch((error) => {
    console.error(error);
    process.exit(1);
  });
}

module.exports = { runJob, checkSheets };
