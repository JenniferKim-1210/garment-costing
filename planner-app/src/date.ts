import type { Repeat } from './types';

const pad = (n: number) => String(n).padStart(2, '0');

export function ymd(d: Date): string {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

export function parseYmd(s: string): Date {
  const [y, m, d] = s.split('-').map(Number);
  return new Date(y, m - 1, d);
}

export function addDays(s: string, n: number): string {
  const d = parseYmd(s);
  d.setDate(d.getDate() + n);
  return ymd(d);
}

export function today(): string {
  return ymd(new Date());
}

export function hm(d: Date): string {
  return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

export function toMinutes(t: string): number {
  const [h, m] = t.split(':').map(Number);
  return h * 60 + m;
}

export function fromMinutes(min: number): string {
  const m = Math.max(0, Math.min(24 * 60 - 1, min));
  return `${pad(Math.floor(m / 60))}:${pad(m % 60)}`;
}

/** 날짜+시간을 로컬 타임스탬프로. 시간이 없으면 오전 9시 기준. */
export function toTimestamp(date: string, time?: string): number {
  const d = parseYmd(date);
  const [h, m] = (time ?? '09:00').split(':').map(Number);
  d.setHours(h, m, 0, 0);
  return d.getTime();
}

/** 월을 더하되, 없는 날짜(예: 2월 31일)는 그 달의 말일로 맞춘다. */
function addMonthsClamped(s: string, n: number): string {
  const d = parseYmd(s);
  const day = d.getDate();
  const target = new Date(d.getFullYear(), d.getMonth() + n, 1);
  const last = new Date(target.getFullYear(), target.getMonth() + 1, 0).getDate();
  target.setDate(Math.min(day, last));
  return ymd(target);
}

/** 반복 규칙에 따른 다음 날짜. */
export function nextOccurrence(date: string, repeat: Repeat): string {
  switch (repeat) {
    case 'daily':
      return addDays(date, 1);
    case 'weekdays': {
      let next = addDays(date, 1);
      while ([0, 6].includes(parseYmd(next).getDay())) next = addDays(next, 1);
      return next;
    }
    case 'weekly':
      return addDays(date, 7);
    case 'monthly':
      return addMonthsClamped(date, 1);
    case 'yearly':
      return addMonthsClamped(date, 12);
    default:
      return date;
  }
}

/** start에서 시작해 repeat 규칙으로 반복되는 항목이 target 날짜에 해당하는지. */
export function occursOn(start: string, repeat: Repeat, target: string): boolean {
  if (target < start) return false;
  if (target === start) return true;
  const s = parseYmd(start);
  const t = parseYmd(target);
  const lastOfTarget = new Date(t.getFullYear(), t.getMonth() + 1, 0).getDate();
  // 31일 반복이 30일까지인 달에서는 말일에 표시되도록 맞춘다.
  const sameDay = t.getDate() === Math.min(s.getDate(), lastOfTarget);
  switch (repeat) {
    case 'daily':
      return true;
    case 'weekdays':
      return t.getDay() >= 1 && t.getDay() <= 5;
    case 'weekly':
      return s.getDay() === t.getDay();
    case 'monthly':
      return sameDay;
    case 'yearly':
      return s.getMonth() === t.getMonth() && sameDay;
    default:
      return false;
  }
}

export function startOfWeek(s: string, weekStart: 0 | 1): string {
  const d = parseYmd(s);
  const diff = (d.getDay() - weekStart + 7) % 7;
  return addDays(s, -diff);
}

/** 달력 그리드용 6주(42일) 날짜 배열. */
export function monthGrid(s: string, weekStart: 0 | 1): string[] {
  const d = parseYmd(s);
  const first = ymd(new Date(d.getFullYear(), d.getMonth(), 1));
  const start = startOfWeek(first, weekStart);
  return Array.from({ length: 42 }, (_, i) => addDays(start, i));
}

export function addMonths(s: string, n: number): string {
  return addMonthsClamped(s, n);
}
