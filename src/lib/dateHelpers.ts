const MONTHS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

function ordinal(n: number): string {
  if (n === 1 || n === 21 || n === 31) return `${n}st`;
  if (n === 2 || n === 22) return `${n}nd`;
  if (n === 3 || n === 23) return `${n}rd`;
  return `${n}th`;
}

// "DM tasks update for 27-04-2026" → "27th April 2026"
export function readableDate(label: string): string {
  const m = label.match(/(\d{1,2})-(\d{1,2})-(\d{4})/);
  if (!m) return label;
  const day = parseInt(m[1]);
  const month = parseInt(m[2]);
  const year = parseInt(m[3]);
  return `${ordinal(day)} ${MONTHS[month - 1]} ${year}`;
}

// "1/5/2026" → "1st May 2026"
export function formatTargetedDate(value: string): string {
  const m = value.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const day = parseInt(m[1]);
    const month = parseInt(m[2]);
    const year = parseInt(m[3]);
    return `${ordinal(day)} ${MONTHS[month - 1]} ${year}`;
  }
  return value.trim();
}
