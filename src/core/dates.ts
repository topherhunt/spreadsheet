// Spreadsheet date serials. In the 1900 system Excel (and Sheets, for
// compatibility) counts 1900-02-29, a day that never existed, so serials from
// 61 on are one day ahead of a plain day count from 1899-12-31.

const DAY_MS = 86400000;
const EPOCH_1900 = Date.UTC(1899, 11, 30);
const EPOCH_1904 = Date.UTC(1904, 0, 1);

/** Serial for a UTC timestamp (ms). */
export function serialFromMs(ms: number, date1904 = false): number {
  if (date1904) return (ms - EPOCH_1904) / DAY_MS;
  const s = (ms - EPOCH_1900) / DAY_MS;
  return s < 61 ? s - 1 : s;
}

/** UTC timestamp (ms) for a serial. */
export function msFromSerial(serial: number, date1904 = false): number {
  if (date1904) return EPOCH_1904 + serial * DAY_MS;
  return EPOCH_1900 + (serial < 61 ? serial + 1 : serial) * DAY_MS;
}

/** Serial for an ISO 8601 date/time as found in <c t="d"> cells, or null if it isn't one. */
export function serialFromIso(iso: string, date1904 = false): number | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?Z?$/.exec(iso);
  if (m) {
    const ms = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] ?? 0), +(m[5] ?? 0), 0) + parseFloat(m[6] ?? '0') * 1000;
    return serialFromMs(ms, date1904);
  }
  const t = /^T?(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?$/.exec(iso);
  if (t) return (+t[1] * 3600 + +t[2] * 60 + parseFloat(t[3] ?? '0')) / 86400;
  return null;
}
