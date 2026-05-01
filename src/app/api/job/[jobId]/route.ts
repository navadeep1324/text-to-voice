import { NextRequest } from 'next/server';
import { jobStore } from '@/lib/jobStore';

export const dynamic = 'force-dynamic';

export async function GET(_req: NextRequest, { params }: { params: { jobId: string } }) {
  const job = jobStore.get(params.jobId);
  if (!job) {
    return Response.json({ error: 'Job not found' }, { status: 404 });
  }
  return Response.json(job);
}
