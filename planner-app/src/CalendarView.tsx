import { useEffect, useRef, useState } from 'react';
import { addDays, addMonths, fromMinutes, hm, monthGrid, occursOn, parseYmd, startOfWeek, toMinutes, today } from './date';
import { fmtDate, fmtTime, t } from './i18n';
import { useStore } from './store';
import { completeTask } from './TasksView';
import type { CalEvent, State, Task } from './types';
import { Icon } from './ui';

type ViewMode = 'month' | 'week' | 'threeDay' | 'day' | 'agenda';
type DayItem = { kind: 'event'; ev: CalEvent } | { kind: 'task'; task: Task };

const MODES: ViewMode[] = ['month', 'week', 'threeDay', 'day', 'agenda'];
const HOUR = 52;

function itemsOn(s: State, day: string): DayItem[] {
  const evs: DayItem[] = s.events.filter((e) => occursOn(e.date, e.repeat, day)).map((ev) => ({ kind: 'event', ev }));
  const tasks: DayItem[] = s.tasks.filter((x) => x.date === day).map((task) => ({ kind: 'task', task }));
  const key = (i: DayItem) => (i.kind === 'event' ? (i.ev.allDay ? '' : i.ev.start) : i.task.time ?? '');
  return [...evs, ...tasks].sort((a, b) => key(a).localeCompare(key(b)));
}

interface Props {
  onEditTask: (task: Task) => void;
  onEditEvent: (ev: CalEvent) => void;
  onAddEvent: (date: string, start?: string) => void;
}

export function CalendarView(p: Props) {
  const s = useStore();
  const [mode, setMode] = useState<ViewMode>('month');
  const [cursor, setCursor] = useState(today());
  const now = today();
  const step = (dir: 1 | -1) => {
    if (mode === 'month') setCursor(addMonths(cursor, dir));
    else setCursor(addDays(cursor, dir * (mode === 'week' ? 7 : mode === 'threeDay' ? 3 : mode === 'day' ? 1 : 14)));
  };
  const days = mode === 'week' ? Array.from({ length: 7 }, (_, i) => addDays(startOfWeek(cursor, s.settings.weekStart), i)) : mode === 'threeDay' ? [0, 1, 2].map((i) => addDays(cursor, i)) : [cursor];

  return (
    <div className="view cal">
      <header className="top">
        <div>
          <p className="eyebrow">{t('calendar')}</p>
          <h1>{fmtDate(parseYmd(cursor), mode === 'day' ? { month: 'long', day: 'numeric', weekday: 'short' } : { year: 'numeric', month: 'long' })}</h1>
        </div>
        <div className="nav-btns">
          <button className="icon-btn" onClick={() => step(-1)} aria-label="previous">
            <Icon name="left" />
          </button>
          <button className="text-btn" onClick={() => setCursor(now)}>
            {t('today')}
          </button>
          <button className="icon-btn" onClick={() => step(1)} aria-label="next">
            <Icon name="right" />
          </button>
        </div>
      </header>
      <div className="seg wide">
        {MODES.map((m) => (
          <button key={m} className={mode === m ? 'on' : ''} onClick={() => setMode(m)}>
            {t(m)}
          </button>
        ))}
      </div>

      {mode === 'month' && <MonthView s={s} cursor={cursor} setCursor={setCursor} {...p} />}
      {(mode === 'week' || mode === 'threeDay' || mode === 'day') && <TimeGrid s={s} days={days} {...p} />}
      {mode === 'agenda' && <Agenda s={s} from={cursor} {...p} />}
    </div>
  );
}

function ItemLine({ item, onEditTask, onEditEvent }: { item: DayItem } & Pick<Props, 'onEditTask' | 'onEditEvent'>) {
  if (item.kind === 'event') {
    const { ev } = item;
    return (
      <button className="line" style={{ ['--c' as string]: ev.color }} onClick={() => onEditEvent(ev)}>
        <span className="line-time">{ev.allDay ? t('allDay') : fmtTime(ev.start)}</span>
        <span className="line-main">
          <b>{ev.title}</b>
          <small>
            {!ev.allDay && `${fmtTime(ev.start)} – ${fmtTime(ev.end)}`}
            {ev.location && ` · ${ev.location}`}
          </small>
        </span>
      </button>
    );
  }
  const { task } = item;
  return (
    <div className="line task-line">
      {task.time && <span className="line-time">{fmtTime(task.time)}</span>}
      <button className={`check p${task.priority} ${task.done ? 'on' : ''}`} role="checkbox" aria-checked={task.done} aria-label={task.title} onClick={() => completeTask(task)}>
        {task.done && <Icon name="check" size={14} />}
      </button>
      <button className={`line-main ${task.done ? 'is-done' : ''}`} onClick={() => onEditTask(task)}>
        <b>{task.title}</b>
      </button>
    </div>
  );
}

function MonthView({ s, cursor, setCursor, onEditTask, onEditEvent, onAddEvent }: Props & { s: State; cursor: string; setCursor: (d: string) => void }) {
  const grid = monthGrid(cursor, s.settings.weekStart);
  const month = cursor.slice(0, 7);
  const now = today();
  const selected = itemsOn(s, cursor);
  return (
    <>
      <div className="month">
        {grid.slice(0, 7).map((d) => (
          <div key={d} className="dow">
            {fmtDate(parseYmd(d), { weekday: 'short' })}
          </div>
        ))}
        {grid.map((d) => {
          const items = itemsOn(s, d).filter((i) => i.kind === 'event' || !i.task.done);
          return (
            <button key={d} className={`cell ${d.slice(0, 7) !== month ? 'out' : ''} ${d === cursor ? 'sel' : ''}`} onClick={() => setCursor(d)} aria-label={fmtDate(parseYmd(d), { dateStyle: 'full' })} aria-pressed={d === cursor}>
              <span className={`num ${d === now ? 'now' : ''}`}>{parseYmd(d).getDate()}</span>
              {items.slice(0, 2).map((i) => (
                <span key={i.kind === 'event' ? i.ev.id : i.task.id} className="pill" style={{ ['--c' as string]: i.kind === 'event' ? i.ev.color : 'var(--muted)' }}>
                  {i.kind === 'event' ? i.ev.title : i.task.title}
                </span>
              ))}
              {items.length > 2 && <span className="more">+{items.length - 2}</span>}
            </button>
          );
        })}
      </div>
      <section className="section">
        <div className="section-head">
          <h2>{fmtDate(parseYmd(cursor), { month: 'long', day: 'numeric', weekday: 'long' })}</h2>
          <button className="icon-btn sm" onClick={() => onAddEvent(cursor)} aria-label={t('addEvent')}>
            <Icon name="plus" size={18} />
          </button>
        </div>
        {selected.length === 0 && <p className="empty sm">{t('emptyDay')}</p>}
        {selected.map((i) => (
          <ItemLine key={i.kind === 'event' ? i.ev.id : i.task.id} item={i} onEditTask={onEditTask} onEditEvent={onEditEvent} />
        ))}
      </section>
    </>
  );
}

interface Block {
  item: DayItem;
  start: number;
  end: number;
  col: number;
  cols: number;
}

/** 겹치는 일정을 나란히 놓기 위한 열 배치. */
export function layout(items: { item: DayItem; start: number; end: number }[]): Block[] {
  const sorted = [...items].sort((a, b) => a.start - b.start || b.end - a.end);
  const out: Block[] = [];
  let cluster: Block[] = [];
  let ends: number[] = [];
  const flush = () => {
    for (const b of cluster) b.cols = ends.length;
    out.push(...cluster);
    cluster = [];
    ends = [];
  };
  for (const it of sorted) {
    if (cluster.length && ends.every((e) => e <= it.start)) flush();
    let col = ends.findIndex((e) => e <= it.start);
    if (col === -1) col = ends.length;
    ends[col] = it.end;
    cluster.push({ ...it, col, cols: 1 });
  }
  flush();
  return out;
}

function TimeGrid({ s, days, onEditTask, onEditEvent, onAddEvent }: Props & { s: State; days: string[] }) {
  const scroller = useRef<HTMLDivElement>(null);
  const now = today();
  const [nowMin, setNowMin] = useState(toMinutes(hm(new Date())));
  useEffect(() => {
    if (scroller.current) scroller.current.scrollTop = Math.max(0, (nowMin / 60 - 1.5) * HOUR);
    const id = setInterval(() => setNowMin(toMinutes(hm(new Date()))), 60000);
    return () => clearInterval(id);
    // 처음 열 때 한 번만 현재 시각 근처로 스크롤한다.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const perDay = days.map((d) => {
    const items = itemsOn(s, d);
    const allDay = items.filter((i) => (i.kind === 'event' ? i.ev.allDay : !i.task.time));
    const timed = items
      .filter((i) => !allDay.includes(i))
      .map((i) => {
        const start = toMinutes(i.kind === 'event' ? i.ev.start : i.task.time!);
        const end = i.kind === 'event' ? Math.max(toMinutes(i.ev.end), start + 20) : start + 30;
        return { item: i, start, end };
      });
    return { day: d, allDay, blocks: layout(timed) };
  });

  return (
    <div className={`grid-wrap n${days.length}`}>
      <div className="grid-head" style={{ gridTemplateColumns: `52px repeat(${days.length}, 1fr)` }}>
        <span />
        {perDay.map(({ day, allDay }) => (
          <div key={day} className="grid-day">
            <span className={`num ${day === now ? 'now' : ''}`}>{parseYmd(day).getDate()}</span>
            <small>{fmtDate(parseYmd(day), { weekday: 'short' })}</small>
            {allDay.map((i) =>
              i.kind === 'event' ? (
                <button key={i.ev.id} className="pill" style={{ ['--c' as string]: i.ev.color }} onClick={() => onEditEvent(i.ev)}>
                  {i.ev.title}
                </button>
              ) : (
                <button key={i.task.id} className={`pill ${i.task.done ? 'is-done' : ''}`} style={{ ['--c' as string]: 'var(--muted)' }} onClick={() => onEditTask(i.task)}>
                  {i.task.title}
                </button>
              ),
            )}
          </div>
        ))}
      </div>
      <div className="grid-scroll" ref={scroller}>
        <div className="grid-body" style={{ gridTemplateColumns: `52px repeat(${days.length}, 1fr)`, height: 24 * HOUR }}>
          <div className="hours">
            {Array.from({ length: 24 }, (_, h) => (
              <span key={h} style={{ top: h * HOUR }}>
                {h > 0 && fmtDate(new Date(2000, 0, 1, h), { hour: 'numeric' })}
              </span>
            ))}
          </div>
          {perDay.map(({ day, blocks }) => (
            <div
              key={day}
              className="col"
              onClick={(e) => {
                if (e.target !== e.currentTarget) return;
                const y = e.clientY - e.currentTarget.getBoundingClientRect().top;
                onAddEvent(day, fromMinutes(Math.floor(y / HOUR) * 60));
              }}
            >
              {day === now && <i className="now-line" style={{ top: (nowMin / 60) * HOUR }} />}
              {blocks.map((b) => {
                const style = { top: (b.start / 60) * HOUR, height: ((b.end - b.start) / 60) * HOUR - 2, left: `${(b.col / b.cols) * 100}%`, width: `calc(${100 / b.cols}% - 2px)` };
                return b.item.kind === 'event' ? (
                  <button key={b.item.ev.id} className="block" style={{ ...style, ['--c' as string]: b.item.ev.color }} onClick={() => onEditEvent((b.item as { ev: CalEvent }).ev)}>
                    <b>{b.item.ev.title}</b>
                    <small>{fmtTime(b.item.ev.start)}</small>
                  </button>
                ) : (
                  <button key={b.item.task.id} className={`block task-block ${b.item.task.done ? 'is-done' : ''}`} style={{ ...style, ['--c' as string]: 'var(--muted)' }} onClick={() => onEditTask((b.item as { task: Task }).task)}>
                    <b>{b.item.task.title}</b>
                  </button>
                );
              })}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function Agenda({ s, from, onEditTask, onEditEvent, onAddEvent }: Props & { s: State; from: string }) {
  const now = today();
  const days = Array.from({ length: 30 }, (_, i) => addDays(from, i))
    .map((d) => ({ d, items: itemsOn(s, d) }))
    .filter((x) => x.items.length > 0 || x.d === now);
  return (
    <div>
      {days.length === 0 && <p className="empty">{t('emptyDay')}</p>}
      {days.map(({ d, items }) => (
        <section key={d} className="section">
          <div className="section-head">
            <h2 className={d === now ? 'accent' : ''}>{fmtDate(parseYmd(d), { month: 'short', day: 'numeric', weekday: 'short' })}</h2>
            <button className="icon-btn sm" onClick={() => onAddEvent(d)} aria-label={t('addEvent')}>
              <Icon name="plus" size={18} />
            </button>
          </div>
          {items.length === 0 && <p className="empty sm">{t('emptyDay')}</p>}
          {items.map((i) => (
            <ItemLine key={i.kind === 'event' ? i.ev.id : i.task.id} item={i} onEditTask={onEditTask} onEditEvent={onEditEvent} />
          ))}
        </section>
      ))}
    </div>
  );
}
