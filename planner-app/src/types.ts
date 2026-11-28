export type Repeat = 'none' | 'daily' | 'weekdays' | 'weekly' | 'monthly' | 'yearly';
export type Lang = 'ko' | 'en';

export interface Subtask {
  id: string;
  title: string;
  done: boolean;
}

export interface Task {
  id: string;
  title: string;
  note: string;
  listId: string;
  done: boolean;
  doneAt?: number;
  /** YYYY-MM-DD. 없으면 "언젠가" */
  date?: string;
  /** HH:mm */
  time?: string;
  /** 마감 몇 분 전에 알릴지. null이면 알림 없음 */
  remind: number | null;
  repeat: Repeat;
  priority: 0 | 1 | 2 | 3;
  subtasks: Subtask[];
  order: number;
  createdAt: number;
}

export interface CalEvent {
  id: string;
  title: string;
  note: string;
  location: string;
  /** 시작 날짜 YYYY-MM-DD (반복 시 첫 날짜) */
  date: string;
  allDay: boolean;
  start: string;
  end: string;
  color: string;
  repeat: Repeat;
  remind: number | null;
}

export interface TaskList {
  id: string;
  name: string;
  color: string;
}

export interface Settings {
  lang: Lang;
  theme: 'system' | 'light' | 'dark';
  weekStart: 0 | 1;
  focusMin: number;
  breakMin: number;
}

export interface FocusLog {
  /** YYYY-MM-DD → 완료한 집중 세션 수 */
  [date: string]: number;
}

export interface State {
  tasks: Task[];
  events: CalEvent[];
  lists: TaskList[];
  settings: Settings;
  focusLog: FocusLog;
}

export const COLORS = ['#0F766E', '#2563EB', '#C2410C', '#B4236B', '#7C3AED', '#4D7C0F', '#A16207', '#475569'];
