export interface JobResult {
  status: 'processing' | 'done' | 'error';
  tab_name?: string;
  date_id?: string;
  filename?: string;
  voice_text?: string;
  audio_url?: string;
  audio_base64?: string;
  error?: string;
}

export const jobStore = new Map<string, JobResult>();
