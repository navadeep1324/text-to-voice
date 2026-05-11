import { NextRequest } from 'next/server';

export const dynamic = 'force-dynamic';
export const runtime = 'nodejs';

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
    const { checkSheets } = require('../../../../../scripts/generate-sharepoint-audio.cjs') as {
      checkSheets: () => Promise<{ today: string; filled: string[]; unfilled: string[] }>;
    };

    const output = await checkSheets();
    return Response.json(output);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return Response.json({ error: message }, { status: 500 });
  }
}
