import * as XLSX from 'xlsx';

export function parseSheetRows(buffer: Buffer, sheetName: string): (string | null)[][] {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  if (!wb.SheetNames.includes(sheetName)) {
    throw new Error(`Sheet '${sheetName}' not found. Available: ${wb.SheetNames.join(', ')}`);
  }
  return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null });
}

export function buildSectionedVoiceText(rows: (string | null)[][]): string {
  const sections: { header: string; sub: string; tasks: string[] }[] = [];
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

    let segment = col1.endsWith('.') ? col1 : col1 + '.';
    if (col2) segment += ` Targeted date ${col2.replace(/\n/g, ' ')}.`;
    if (col3) segment += ` Status: ${col3}.`;

    if (sections.length === 0) sections.push({ header: '', sub: '', tasks: [] });
    sections[sections.length - 1].tasks.push(segment);
  }

  const parts: string[] = [];
  if (dateIntro) parts.push(`Updates for ${dateIntro}.`);
  for (const sec of sections) {
    if (sec.tasks.length === 0) continue;
    if (sec.header) parts.push(`${sec.header}.`);
    if (sec.sub) parts.push(`${sec.sub}.`);
    parts.push(...sec.tasks);
  }

  return parts.join(' ');
}
