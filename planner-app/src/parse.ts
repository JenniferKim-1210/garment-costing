import { addDays, ymd } from './date';

export interface Parsed {
  title: string;
  date?: string;
  time?: string;
}

const KO_DAYS = ['일', '월', '화', '수', '목', '금', '토'];
const EN_DAYS = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
const pad = (n: number) => String(n).padStart(2, '0');

/**
 * "내일 오후 3시 회의", "금요일 장보기", "tomorrow 9am standup" 같은 문장에서
 * 날짜와 시간을 뽑아내고 나머지를 제목으로 돌려준다.
 */
export function parseQuick(input: string, now: Date = new Date()): Parsed {
  let text = input;
  const base = ymd(now);
  let date: string | undefined;
  let time: string | undefined;

  const take = (re: RegExp, fn: (m: RegExpMatchArray) => boolean | void) => {
    const m = text.match(re);
    if (m && fn(m) !== false) text = text.replace(m[0], ' ');
  };

  // 날짜
  take(/(\d{1,2})월\s*(\d{1,2})일/, (m) => {
    const mo = +m[1], d = +m[2];
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return false;
    let y = now.getFullYear();
    if (`${y}-${pad(mo)}-${pad(d)}` < base) y += 1;
    date = `${y}-${pad(mo)}-${pad(d)}`;
  });
  if (!date) take(/오늘|\btoday\b/i, () => { date = base; });
  if (!date) take(/모레/, () => { date = addDays(base, 2); });
  if (!date) take(/내일|\btomorrow\b/i, () => { date = addDays(base, 1); });
  if (!date) take(/다음\s*주|\bnext week\b/i, () => { date = addDays(base, 7); });
  if (!date) take(/([일월화수목금토])요일/, (m) => {
    const diff = (KO_DAYS.indexOf(m[1]) - now.getDay() + 7) % 7;
    date = addDays(base, diff);
  });
  if (!date) take(new RegExp(`\\b(${EN_DAYS.join('|')})\\b`, 'i'), (m) => {
    const diff = (EN_DAYS.indexOf(m[1].toLowerCase()) - now.getDay() + 7) % 7;
    date = addDays(base, diff);
  });

  // 시간
  take(/(오전|오후)?\s*(\d{1,2})시\s*(?:(\d{1,2})분|(반))?/, (m) => {
    let h = +m[2];
    const min = m[4] ? 30 : m[3] ? +m[3] : 0;
    if (h > 24 || min > 59) return false;
    if (m[1] === '오후' && h < 12) h += 12;
    else if (m[1] === '오전' && h === 12) h = 0;
    // 오전/오후를 안 쓴 1~6시는 보통 오후를 뜻한다.
    else if (!m[1] && h >= 1 && h <= 6) h += 12;
    time = `${pad(h % 24)}:${pad(min)}`;
  });
  if (!time) take(/\b(\d{1,2})(?::(\d{2}))?\s*(am|pm)\b/i, (m) => {
    let h = +m[1];
    const min = m[2] ? +m[2] : 0;
    if (h < 1 || h > 12 || min > 59) return false;
    const pm = m[3].toLowerCase() === 'pm';
    if (pm && h < 12) h += 12;
    if (!pm && h === 12) h = 0;
    time = `${pad(h)}:${pad(min)}`;
  });
  if (!time) take(/\b([01]?\d|2[0-3]):([0-5]\d)\b/, (m) => { time = `${pad(+m[1])}:${m[2]}`; });

  // 시간만 있으면 오늘로 본다.
  if (time && !date) date = base;

  const title = text.replace(/\s+/g, ' ').replace(/^(에|at)\s+/i, '').trim();
  return { title: title || input.trim(), date, time };
}
