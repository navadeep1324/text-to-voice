import Papa from 'papaparse';
import { readableDate, formatTargetedDate } from './dateHelpers';

interface GSConfig {
  sheet_id: string;
  tabs: string[];
  date_column?: string;
  tasks_column?: string;
  targeted_date_column?: string;
  remarks_column?: string;
}

export type SheetItem = [string, string, string]; // [tab_name, date_id, voice_text]

export class GoogleSheetReader {
  private sheetId: string;
  private tabs: string[];

  constructor(config: GSConfig) {
    this.sheetId = config.sheet_id;
    this.tabs = config.tabs || [];
  }

  async readAllTabs(): Promise<SheetItem[]> {
    const results: SheetItem[] = [];
    for (const tab of this.tabs) {
      try {
        const items = await this.readTab(tab);
        results.push(...items);
      } catch (e) {
        console.error(`Failed to read tab '${tab}':`, e);
      }
    }
    return results;
  }

  private async readTab(tabName: string): Promise<SheetItem[]> {
    const url = `https://docs.google.com/spreadsheets/d/${this.sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(tabName)}`;
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`HTTP ${resp.status} fetching tab '${tabName}'`);

    const csvText = await resp.text();
    const parsed = Papa.parse<string[]>(csvText, { skipEmptyLines: true });
    const rows = parsed.data;

    if (rows.length === 0) return [];

    // Row 0 has merged header+data: "Date <value>", "Tasks <value>"
    rows[0][0] = (rows[0][0] || '').replace(/^Date\s+/, '').trim();
    rows[0][1] = (rows[0][1] || '').replace(/^Tasks\s+/, '').trim();
    rows[0][2] = '';
    rows[0][3] = '';

    console.log(`Downloaded tab '${tabName}': ${rows.length} rows`);
    return this.buildSummaries(rows, tabName);
  }

  private buildSummaries(rows: string[][], tabName: string): SheetItem[] {
    // Forward-fill the date column
    let currentDate = '';
    const filled = rows.map(row => {
      if (row[0]?.trim()) currentDate = row[0].trim();
      return {
        date: currentDate,
        task: (row[1] || '').trim(),
        targetedDate: (row[2] || '').trim(),
        remarks: (row[3] || '').trim(),
      };
    });

    // Group by date preserving insertion order
    const groups = new Map<string, typeof filled>();
    for (const row of filled) {
      if (!groups.has(row.date)) groups.set(row.date, []);
      groups.get(row.date)!.push(row);
    }

    const summaries: SheetItem[] = [];
    for (const [dateLabel, group] of groups) {
      const parts: string[] = [];
      for (const row of group) {
        if (!row.task) continue;

        let segment = row.task.endsWith('.') ? row.task : row.task + '.';
        if (row.targetedDate) {
          segment += ` Targeted date ${formatTargetedDate(row.targetedDate)}.`;
        }
        if (row.remarks) {
          segment += ` Remarks ${row.remarks}.`;
        }
        parts.push(segment);
      }

      if (parts.length === 0) continue;

      const readable = readableDate(dateLabel);
      const voiceText = parts.join(' ');
      const dateId = dateLabel.replace(/[^\w]/g, '_');
      summaries.push([tabName, dateId, voiceText]);
      console.log(`Tab '${tabName}' / '${dateLabel}': ${parts.length} task line(s)`);
    }

    return summaries;
  }
}
