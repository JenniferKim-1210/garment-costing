import { describe, expect, it } from 'vitest';
import { addDays, monthGrid, nextOccurrence, occursOn, startOfWeek } from './date';
import { parseQuick } from './parse';

describe('반복 규칙', () => {
  it('다음 회차를 계산한다', () => {
    expect(nextOccurrence('2026-11-28', 'daily')).toBe('2026-11-29');
    expect(nextOccurrence('2026-11-27', 'weekdays')).toBe('2026-11-30'); // 금 → 월
    expect(nextOccurrence('2026-11-28', 'weekly')).toBe('2026-12-05');
    expect(nextOccurrence('2026-01-31', 'monthly')).toBe('2026-02-28'); // 말일로 맞춤
    expect(nextOccurrence('2024-02-29', 'yearly')).toBe('2025-02-28');
    expect(nextOccurrence('2026-12-31', 'daily')).toBe('2027-01-01');
  });

  it('특정 날짜에 해당하는지 판단한다', () => {
    expect(occursOn('2026-11-28', 'none', '2026-11-28')).toBe(true);
    expect(occursOn('2026-11-28', 'none', '2026-11-29')).toBe(false);
    expect(occursOn('2026-11-28', 'daily', '2026-11-27')).toBe(false); // 시작 전
    expect(occursOn('2026-11-27', 'weekdays', '2026-11-28')).toBe(false); // 토요일
    expect(occursOn('2026-11-27', 'weekdays', '2026-11-30')).toBe(true);
    expect(occursOn('2026-11-28', 'weekly', '2026-12-12')).toBe(true);
    expect(occursOn('2026-01-31', 'monthly', '2026-02-28')).toBe(true);
    expect(occursOn('2026-01-31', 'monthly', '2026-03-30')).toBe(false);
    expect(occursOn('2026-03-15', 'yearly', '2027-03-15')).toBe(true);
  });
});

describe('달력 계산', () => {
  it('주 시작일과 월 그리드를 만든다', () => {
    expect(startOfWeek('2026-11-28', 0)).toBe('2026-11-22');
    expect(startOfWeek('2026-11-28', 1)).toBe('2026-11-23');
    const grid = monthGrid('2026-11-28', 0);
    expect(grid).toHaveLength(42);
    expect(grid[0]).toBe('2026-11-01');
    expect(addDays('2026-02-28', 1)).toBe('2026-03-01');
  });
});

describe('빠른 입력 해석', () => {
  const now = new Date(2026, 10, 28, 10, 0); // 2026-11-28 토요일

  it('한국어 날짜와 시간을 읽는다', () => {
    expect(parseQuick('내일 오후 3시 회의', now)).toEqual({ title: '회의', date: '2026-11-29', time: '15:00' });
    expect(parseQuick('모레 장보기', now)).toEqual({ title: '장보기', date: '2026-11-30', time: undefined });
    expect(parseQuick('금요일 9시 반 치과', now)).toEqual({ title: '치과', date: '2026-12-04', time: '09:30' });
    expect(parseQuick('3시 미팅', now)).toEqual({ title: '미팅', date: '2026-11-28', time: '15:00' });
    expect(parseQuick('12월 25일 선물 사기', now)).toEqual({ title: '선물 사기', date: '2026-12-25', time: undefined });
    expect(parseQuick('1월 5일 신년회', now).date).toBe('2027-01-05'); // 지난 날짜는 내년으로
    expect(parseQuick('오전 12시 10분 알람', now).time).toBe('00:10');
  });

  it('영어 날짜와 시간을 읽는다', () => {
    expect(parseQuick('Standup tomorrow 9am', now)).toEqual({ title: 'Standup', date: '2026-11-29', time: '09:00' });
    expect(parseQuick('Call mom monday 18:30', now)).toEqual({ title: 'Call mom', date: '2026-11-30', time: '18:30' });
    expect(parseQuick('Lunch 12pm', now).time).toBe('12:00');
  });

  it('날짜가 없으면 제목을 그대로 둔다', () => {
    expect(parseQuick('우유 사기', now)).toEqual({ title: '우유 사기', date: undefined, time: undefined });
    expect(parseQuick('  ', now).title).toBe('');
  });
});
