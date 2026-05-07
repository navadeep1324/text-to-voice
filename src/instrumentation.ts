export async function register() {
  // Only run in Node.js runtime (not Edge), and only once per process
  if (process.env.NEXT_RUNTIME === 'edge') return;
  if (process.env.ENABLE_INTERNAL_SCHEDULER !== 'true') {
    console.log('[Scheduler] Internal scheduler disabled. Use hosting cron for scheduled audio generation.');
    return;
  }

  const { default: cron } = await import('node-cron');
  const { loadConfig } = await import('@/lib/config');
  const { generateSharePointAudio } = await import('@/app/api/generate-sharepoint/route');

  const config = loadConfig();
  const schedule = config.schedule ?? { cron: '0 9 * * 1-5', timezone: 'Asia/Kolkata' };

  if (!config.sharepoint) {
    console.log('[Scheduler] SharePoint not configured — daily job skipped');
    return;
  }

  console.log(`[Scheduler] Daily audio job scheduled: "${schedule.cron}" (${schedule.timezone})`);

  cron.schedule(
    schedule.cron,
    async () => {
      console.log('[Scheduler] Running daily SharePoint audio generation...');
      try {
        const results = await generateSharePointAudio('http', 'localhost:8000');
        console.log(`[Scheduler] Done. Generated ${results.length} audio file(s):`, results.map((r) => r.filename));
      } catch (e) {
        console.error('[Scheduler] Daily job failed:', e instanceof Error ? e.message : e);
      }
    },
    { timezone: schedule.timezone }
  );
}
