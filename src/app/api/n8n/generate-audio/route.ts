import { NextRequest } from 'next/server';

export const dynamic = 'force-dynamic';
export const runtime = 'nodejs';

type AudioJobResult = {
  result: string;
  workbook: string;
  today: string;
  results: Array<{
    sheet: string;
    filename: string;
    bytes: number;
    contentType?: string;
    base64?: string;
  }>;
  skipped: Array<{
    sheet: string;
    reason: string;
  }>;
};

function isAuthorized(request: NextRequest) {
  const expectedKey = process.env.N8N_API_KEY;
  if (!expectedKey) return false;

  const authHeader = request.headers.get('authorization') || '';
  const bearerToken = authHeader.startsWith('Bearer ') ? authHeader.slice('Bearer '.length) : '';
  const headerKey = request.headers.get('x-api-key') || '';
  return bearerToken === expectedKey || headerKey === expectedKey;
}

export async function POST(request: NextRequest) {
  if (!isAuthorized(request)) {
    return Response.json({ error: 'Unauthorized' }, { status: 401 });
  }

  try {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const { runJob } = require('../../../../../scripts/generate-sharepoint-audio.cjs') as {
      runJob: (options: {
        persistAudio: boolean;
        notifyTeams: boolean;
        includeBase64: boolean;
      }) => Promise<AudioJobResult>;
    };

    const output = await runJob({
      persistAudio: false,
      notifyTeams: false,
      includeBase64: true,
    });

    return Response.json(output);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return Response.json({ error: message }, { status: 500 });
  }
}
