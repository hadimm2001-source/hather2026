import { Config, TimeObj } from '../types';

export function norm(s: any): string {
  return String(s ?? '').trim();
}

export function normKey(s: any): string {
  return norm(s)
    .replace(/[إأآا]/g, 'ا')
    .replace(/ى/g, 'ي')
    .replace(/ة/g, 'ه')
    .replace(/\s+/g, ' ')
    .trim();
}

export function getVal(o: any, names: string[]): string {
  for (const n of names) {
    const k = normKey(n);
    if (Object.prototype.hasOwnProperty.call(o, k)) {
      return o[k];
    }
  }
  return '';
}

export function pad(n: number): string {
  return String(n).padStart(2, '0');
}

export function parseTime(v: any): TimeObj | null {
  if (v == null || v === '') return null;
  if (typeof v === 'number') {
    let m = Math.round(v * 24 * 60);
    return {
      h: Math.floor(m / 60) % 24,
      m: m % 60,
      total: m,
      text: `${pad(Math.floor(m / 60) % 24)}:${pad(m % 60)}`,
    };
  }
  let s = norm(v);
  if (s === '-' || s === '—' || s === 'لا يوجد' || s === 'بلا بصمة') {
    return { text: '-', total: 0 } as any; // special block or empty
  }
  let m = s.match(/(\d{1,2})[:٫.](\d{1,2})/);
  if (!m) return null;
  let h = +m[1];
  let mi = +m[2];
  return {
    h,
    m: mi,
    total: h * 60 + mi,
    text: `${pad(h)}:${pad(mi)}`,
  };
}

export function parseDate(v: any): Date | null {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === 'number') {
    // Excel base date is Dec 30, 1899 due to 1900 leap year bug
    return new Date(Math.round((v - 25569) * 86400 * 1000));
  }
  let s = norm(v).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d).toString());
  let m = s.match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})/);
  if (!m) {
    // try YYYY-MM-DD
    let m2 = s.match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    if (m2) {
      return new Date(+m2[1], +m2[2] - 1, +m2[3]);
    }
    return null;
  }
  let d = +m[1];
  let mo = +m[2] - 1;
  let y = +m[3];
  if (y < 100) y += 2000;
  return new Date(y, mo, d);
}

export function dateKey(d: Date): string {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

export function arDate(d: Date): string {
  return `${pad(d.getDate())}-${pad(d.getMonth() + 1)}-${d.getFullYear()}`;
}

export function monthKey(d: Date): string {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
}

export function minsAfterStart(t: TimeObj | null, startStr: string): number {
  const st = parseTime(startStr);
  if (!t || t.total == null || !st) return 0;
  return Math.max(0, t.total - st.total);
}

export function expectedCheckout(tin: TimeObj | null, workHours: number): number | null {
  if (!tin || tin.total == null) return null;
  return tin.total + (Number(workHours) || 7) * 60;
}

export function checkoutState(
  status: string,
  tin: TimeObj | null,
  tout: TimeObj | null,
  autoCheckoutStr: string,
  workHours: number
): string {
  const st = norm(status);
  if (!tout || tout.total == null || tout.text === '-') {
    if (st.includes('غير مكتملة') || (tin && tin.total != null)) return 'لا توجد بصمة انصراف';
    return '';
  }
  
  if (tout.total >= 15 * 60) {
    return 'انصراف في غير وقت الفعلي';
  }

  const auto = parseTime(autoCheckoutStr) || { total: 16 * 60 };
  if (st.includes('انصراف تلقائي') || tout.total >= auto.total) {
    return 'انصراف تلقائي';
  }
  const exp = expectedCheckout(tin, workHours);
  if (exp != null && tout.total < exp) {
    return 'انصراف مبكر';
  }
  return 'مكتمل';
}

export function earlyCheckoutMins(
  tin: TimeObj | null,
  tout: TimeObj | null,
  status: string,
  autoCheckoutStr: string,
  workHours: number
): number {
  if (!tin || tin.total == null || !tout || tout.total == null || tout.text === '-') return 0;
  const state = checkoutState(status, tin, tout, autoCheckoutStr, workHours);
  if (state === 'انصراف تلقائي' || state === 'انصراف في غير وقت الفعلي') return 0;
  const exp = expectedCheckout(tin, workHours);
  return exp == null ? 0 : Math.max(0, exp - tout.total);
}

export function isWeekend(d: Date, weekend: number[]): boolean {
  return weekend.includes(d.getDay());
}

export function parseHolidayDate(s: string): Date | null {
  s = norm(s).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d).toString());
  let m = s.match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})/);
  if (m) {
    let y = +m[3];
    if (y < 100) y += 2000;
    return new Date(y, +m[2] - 1, +m[1]);
  }
  return null;
}

export function addHolidayRange(set: Set<string>, a: Date | null, b: Date | null) {
  if (!a) return;
  if (!b) b = a;
  const start = a <= b ? a : b;
  const end = a <= b ? b : a;
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    set.add(dateKey(d));
  }
}

export function holidaySet(holidaysStr: string): Set<string> {
  const set = new Set<string>();
  const lines = String(holidaysStr || '')
    .split(/[\n,،]+/)
    .map(x => x.trim())
    .filter(Boolean);
  for (const line of lines) {
    const parts = line.match(/\d{4}[-\/]\d{1,2}[-\/]\d{1,2}|\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}/g) || [];
    if (parts.length >= 2) {
      addHolidayRange(set, parseHolidayDate(parts[0]), parseHolidayDate(parts[1]));
    } else if (parts.length === 1) {
      addHolidayRange(set, parseHolidayDate(parts[0]), null);
    }
  }
  return set;
}

export function isHoliday(d: Date, holidaysStr: string): boolean {
  return holidaySet(holidaysStr).has(dateKey(d));
}

export function nonWorkReason(d: Date, weekend: number[], holidaysStr: string): string {
  if (isWeekend(d, weekend)) return 'عطلة أسبوعية';
  if (isHoliday(d, holidaysStr)) return 'إجازة رسمية';
  return '';
}

export function dayName(d: Date): string {
  return ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'][d.getDay()];
}
