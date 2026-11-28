import { useSyncExternalStore } from 'react';
import { addDays, nextOccurrence, today } from './date';
import { tFor } from './i18n';
import type { CalEvent, Lang, State, Task, TaskList } from './types';

const KEY = 'haru.v1';

export function uid(): string {
  return typeof crypto !== 'undefined' && crypto.randomUUID
    ? crypto.randomUUID()
    : Math.random().toString(36).slice(2) + Date.now().toString(36);
}

function initial(): State {
  const lang: Lang = navigator.language?.toLowerCase().startsWith('ko') ? 'ko' : 'en';
  const lists: TaskList[] = [
    { id: 'personal', name: tFor(lang, 'defPersonal'), color: '#0F766E' },
    { id: 'work', name: tFor(lang, 'defWork'), color: '#2563EB' },
    { id: 'shopping', name: tFor(lang, 'defShopping'), color: '#C2410C' },
  ];
  const base = { note: '', done: false, remind: null, repeat: 'none' as const, priority: 0 as const, subtasks: [], createdAt: Date.now() };
  return {
    lists,
    events: [],
    tasks: [
      { ...base, id: uid(), title: tFor(lang, 'sample1'), listId: 'personal', date: today(), order: 0 },
      { ...base, id: uid(), title: tFor(lang, 'sample2'), listId: 'personal', date: today(), order: 1 },
      { ...base, id: uid(), title: tFor(lang, 'sample3'), listId: 'shopping', date: addDays(today(), 1), order: 2 },
    ],
    settings: { lang, theme: 'system', weekStart: 0, focusMin: 25, breakMin: 5 },
    focusLog: {},
  };
}

/** 저장된 JSON이나 가져온 백업을 검증해 State로 만든다. 형식이 다르면 null. */
export function normalize(raw: unknown): State | null {
  if (!raw || typeof raw !== 'object') return null;
  const r = raw as Partial<State>;
  if (!Array.isArray(r.tasks) || !Array.isArray(r.events) || !Array.isArray(r.lists) || r.lists.length === 0) return null;
  const def = initial();
  return {
    tasks: r.tasks.filter((x) => x && typeof x.id === 'string' && typeof x.title === 'string').map((x) => ({
      ...x,
      note: x.note ?? '',
      remind: x.remind ?? null,
      repeat: x.repeat ?? 'none',
      priority: x.priority ?? 0,
      subtasks: Array.isArray(x.subtasks) ? x.subtasks : [],
      order: typeof x.order === 'number' ? x.order : 0,
    })),
    events: r.events.filter((x) => x && typeof x.id === 'string' && typeof x.date === 'string'),
    lists: r.lists.filter((x) => x && typeof x.id === 'string'),
    settings: { ...def.settings, ...(r.settings ?? {}) },
    focusLog: r.focusLog && typeof r.focusLog === 'object' ? r.focusLog : {},
  };
}

function load(): State {
  try {
    const raw = localStorage.getItem(KEY);
    if (raw) {
      const s = normalize(JSON.parse(raw));
      if (s) return s;
    }
  } catch {
    // 손상된 저장 데이터는 무시하고 새로 시작한다.
  }
  return initial();
}

let state: State = load();
const subs = new Set<() => void>();

export function getState(): State {
  return state;
}

export function setState(fn: (s: State) => State) {
  state = fn(state);
  try {
    localStorage.setItem(KEY, JSON.stringify(state));
  } catch {
    // 저장 공간이 가득 찬 경우에도 화면은 계속 동작하게 둔다.
  }
  subs.forEach((f) => f());
}

export function subscribe(f: () => void) {
  subs.add(f);
  return () => {
    subs.delete(f);
  };
}

export function useStore(): State {
  return useSyncExternalStore(subscribe, getState);
}

// ---------- 할 일 ----------

export function newTask(partial: Partial<Task>): Task {
  const maxOrder = state.tasks.reduce((m, x) => Math.max(m, x.order), 0);
  return {
    id: uid(),
    title: '',
    note: '',
    listId: state.lists[0].id,
    done: false,
    remind: null,
    repeat: 'none',
    priority: 0,
    subtasks: [],
    order: maxOrder + 1,
    createdAt: Date.now(),
    ...partial,
  };
}

export function saveTask(task: Task) {
  setState((s) => ({
    ...s,
    tasks: s.tasks.some((x) => x.id === task.id) ? s.tasks.map((x) => (x.id === task.id ? task : x)) : [...s.tasks, task],
  }));
}

export function deleteTask(id: string) {
  setState((s) => ({ ...s, tasks: s.tasks.filter((x) => x.id !== id) }));
}

/**
 * 완료 처리. 반복되는 할 일은 완료 대신 다음 날짜로 넘긴다.
 * 넘어간 경우 다음 날짜를 돌려준다.
 */
export function toggleTask(id: string): string | undefined {
  const task = state.tasks.find((x) => x.id === id);
  if (!task) return;
  if (!task.done && task.repeat !== 'none' && task.date) {
    // 오래 밀린 반복 일정도 과거가 아닌 다음 회차로 넘어가게 한다.
    let next = nextOccurrence(task.date, task.repeat);
    const now = today();
    while (next < now) next = nextOccurrence(next, task.repeat);
    saveTask({ ...task, date: next, subtasks: task.subtasks.map((x) => ({ ...x, done: false })) });
    return next;
  }
  saveTask({ ...task, done: !task.done, doneAt: task.done ? undefined : Date.now() });
}

export function clearCompleted(listId?: string) {
  setState((s) => ({ ...s, tasks: s.tasks.filter((x) => !(x.done && (!listId || x.listId === listId))) }));
}

export type SectionKey = 'overdue' | 'today' | 'tomorrow' | 'upcoming' | 'someday';

export function sectionOf(task: Task, now = today()): SectionKey {
  if (!task.date) return 'someday';
  if (task.date < now) return 'overdue';
  if (task.date === now) return 'today';
  if (task.date === addDays(now, 1)) return 'tomorrow';
  return 'upcoming';
}

/** 끌어서 놓기: 대상 할 일의 앞/뒤로 옮기고, 섹션이 바뀌면 날짜도 바꾼다. */
export function moveTask(id: string, targetId: string | null, before: boolean, section: SectionKey) {
  const now = today();
  setState((s) => {
    const task = s.tasks.find((x) => x.id === id);
    if (!task) return s;
    const target = targetId ? s.tasks.find((x) => x.id === targetId) : undefined;
    let date = task.date;
    if (section === 'today') date = now;
    else if (section === 'tomorrow') date = addDays(now, 1);
    else if (section === 'someday') date = undefined;
    else if (section === 'upcoming') date = target?.date ?? (sectionOf(task, now) === 'upcoming' ? task.date : addDays(now, 2));
    else if (sectionOf(task, now) !== 'overdue') return s; // 지난 일 섹션으로는 옮기지 않는다.
    const order = target ? target.order + (before ? -0.5 : 0.5) : Number.MAX_SAFE_INTEGER;
    const moved = { ...task, date, time: date ? task.time : undefined, order };
    const sorted = s.tasks.map((x) => (x.id === id ? moved : x)).sort((a, b) => a.order - b.order);
    return { ...s, tasks: sorted.map((x, i) => ({ ...x, order: i })) };
  });
}

// ---------- 일정 ----------

export function saveEvent(ev: CalEvent) {
  setState((s) => ({
    ...s,
    events: s.events.some((x) => x.id === ev.id) ? s.events.map((x) => (x.id === ev.id ? ev : x)) : [...s.events, ev],
  }));
}

export function deleteEvent(id: string) {
  setState((s) => ({ ...s, events: s.events.filter((x) => x.id !== id) }));
}

// ---------- 목록 ----------

export function saveList(list: TaskList) {
  setState((s) => ({
    ...s,
    lists: s.lists.some((x) => x.id === list.id) ? s.lists.map((x) => (x.id === list.id ? list : x)) : [...s.lists, list],
  }));
}

export function deleteList(id: string) {
  setState((s) => {
    if (s.lists.length <= 1) return s;
    const lists = s.lists.filter((x) => x.id !== id);
    return { ...s, lists, tasks: s.tasks.map((x) => (x.listId === id ? { ...x, listId: lists[0].id } : x)) };
  });
}

export function wipe() {
  localStorage.removeItem(KEY);
  state = { ...initial(), tasks: [] };
  setState((s) => s);
}
