import { Capacitor } from '@capacitor/core';
import { addDays, occursOn, today, toTimestamp } from './date';
import { fmtTime, t } from './i18n';
import type { State } from './types';

export interface Reminder {
  key: string;
  title: string;
  body: string;
  at: number;
}

/** from 이후 horizonDays 일 안에 울릴 알림 목록. */
export function computeReminders(state: State, from: number, horizonDays = 30): Reminder[] {
  const out: Reminder[] = [];
  for (const task of state.tasks) {
    if (task.done || task.remind === null || !task.date) continue;
    const at = toTimestamp(task.date, task.time) - task.remind * 60000;
    if (at >= from) out.push({ key: `t:${task.id}:${at}`, title: task.title, body: task.time ? fmtTime(task.time) : t('task'), at });
  }
  const start = today();
  for (const ev of state.events) {
    if (ev.remind === null) continue;
    // 하루 전 알림까지 잡히도록 하루 여유를 두고 훑는다.
    for (let i = 0; i <= horizonDays + 1; i++) {
      const day = addDays(start, i);
      if (!occursOn(ev.date, ev.repeat, day)) continue;
      const at = toTimestamp(day, ev.allDay ? undefined : ev.start) - ev.remind * 60000;
      if (at >= from) {
        const when = ev.allDay ? t('allDay') : `${fmtTime(ev.start)} – ${fmtTime(ev.end)}`;
        out.push({ key: `e:${ev.id}:${at}`, title: ev.title, body: ev.location ? `${when} · ${ev.location}` : when, at });
      }
      if (ev.repeat === 'none') break;
    }
  }
  return out.sort((a, b) => a.at - b.at);
}

export const isNative = Capacitor.isNativePlatform();

export type Permission = 'granted' | 'denied' | 'default' | 'unsupported';

export async function getPermission(): Promise<Permission> {
  if (isNative) {
    const { LocalNotifications } = await import('@capacitor/local-notifications');
    const p = await LocalNotifications.checkPermissions();
    return p.display === 'granted' ? 'granted' : p.display === 'denied' ? 'denied' : 'default';
  }
  if (!('Notification' in window)) return 'unsupported';
  return Notification.permission;
}

export async function requestPermission(): Promise<Permission> {
  if (isNative) {
    const { LocalNotifications } = await import('@capacitor/local-notifications');
    const p = await LocalNotifications.requestPermissions();
    return p.display === 'granted' ? 'granted' : 'denied';
  }
  if (!('Notification' in window)) return 'unsupported';
  return Notification.requestPermission();
}

export async function showNow(title: string, body: string) {
  if (isNative) {
    const { LocalNotifications } = await import('@capacitor/local-notifications');
    await LocalNotifications.schedule({ notifications: [{ id: Date.now() % 2147483647, title, body }] });
    return;
  }
  if (!('Notification' in window) || Notification.permission !== 'granted') return;
  try {
    const reg = await navigator.serviceWorker?.getRegistration();
    if (reg) await reg.showNotification(title, { body, icon: './icon-192.png' });
    else new Notification(title, { body, icon: './icon-192.png' });
  } catch {
    // 알림 표시에 실패해도 앱 안의 토스트로는 계속 알려 준다.
  }
}

function hash(s: string): number {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0;
  return Math.abs(h) % 2147483647;
}

/** 네이티브: 예약된 알림을 현재 상태에 맞게 다시 건다(앱이 꺼져 있어도 울린다). */
async function syncNative(state: State) {
  const { LocalNotifications } = await import('@capacitor/local-notifications');
  if ((await LocalNotifications.checkPermissions()).display !== 'granted') return;
  const pending = await LocalNotifications.getPending();
  if (pending.notifications.length) await LocalNotifications.cancel({ notifications: pending.notifications.map((n) => ({ id: n.id })) });
  // 안드로이드의 예약 알림 개수 제한을 넘지 않도록 가까운 것부터 60개만 건다.
  const next = computeReminders(state, Date.now()).slice(0, 60);
  if (next.length) {
    await LocalNotifications.schedule({
      notifications: next.map((r) => ({ id: hash(r.key), title: r.title, body: r.body, schedule: { at: new Date(r.at), allowWhileIdle: true } })),
    });
  }
}

const FIRED_KEY = 'haru.fired';

/** 웹: 앱이 열려 있는 동안 주기적으로 확인해 시간이 된 알림을 띄운다. */
function tickWeb(state: State, onFire: (r: Reminder) => void) {
  const now = Date.now();
  let fired: Record<string, number> = {};
  try {
    fired = JSON.parse(localStorage.getItem(FIRED_KEY) ?? '{}');
  } catch {
    fired = {};
  }
  // 앱을 잠깐 닫았다 연 경우를 위해 최근 10분 안에 지난 알림까지 챙긴다.
  const due = computeReminders(state, now - 10 * 60000).filter((r) => r.at <= now && !fired[r.key]);
  if (!due.length) return;
  for (const r of due) {
    fired[r.key] = now;
    onFire(r);
    void showNow(r.title, r.body);
  }
  for (const k of Object.keys(fired)) if (fired[k] < now - 86400000) delete fired[k];
  localStorage.setItem(FIRED_KEY, JSON.stringify(fired));
}

export function startReminders(getState: () => State, subscribe: (f: () => void) => () => void, onFire: (r: Reminder) => void) {
  if (isNative) {
    let timer: ReturnType<typeof setTimeout> | undefined;
    const sync = () => {
      clearTimeout(timer);
      timer = setTimeout(() => void syncNative(getState()).catch(() => undefined), 500);
    };
    sync();
    const unsub = subscribe(sync);
    return () => {
      clearTimeout(timer);
      unsub();
    };
  }
  const tick = () => tickWeb(getState(), onFire);
  tick();
  const id = setInterval(tick, 20000);
  document.addEventListener('visibilitychange', tick);
  return () => {
    clearInterval(id);
    document.removeEventListener('visibilitychange', tick);
  };
}
