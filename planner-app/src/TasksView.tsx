import { useRef, useState } from 'react';
import type { PointerEvent as RPointerEvent } from 'react';
import { addDays, parseYmd, today } from './date';
import { fmtDate, fmtTime, t } from './i18n';
import { clearCompleted, deleteTask, moveTask, saveTask, sectionOf, toggleTask, useStore } from './store';
import type { SectionKey } from './store';
import type { Task, TaskList } from './types';
import { Icon, shareText, toast } from './ui';

interface DragState {
  id: string;
  dy: number;
  target: string | null;
  before: boolean;
  section: SectionKey | null;
}

const SECTIONS: SectionKey[] = ['overdue', 'today', 'tomorrow', 'upcoming', 'someday'];

export function completeTask(task: Task) {
  const next = toggleTask(task.id);
  if (next) toast(`${t('nextOn')}: ${fmtDate(parseYmd(next), { month: 'short', day: 'numeric', weekday: 'short' })}`);
}

export function removeTask(task: Task) {
  deleteTask(task.id);
  toast(t('deleted'), t('undo'), () => saveTask(task));
}

export function TasksView({ onEdit, onAdd }: { onEdit: (task: Task) => void; onAdd: (date: string | undefined, listId?: string) => void }) {
  const s = useStore();
  const [listId, setListId] = useState<string | null>(null);
  const [showDone, setShowDone] = useState(false);
  const [drag, setDrag] = useState<DragState | null>(null);
  const startY = useRef(0);
  const now = today();

  const visible = s.tasks.filter((x) => !listId || x.listId === listId);
  const open = visible.filter((x) => !x.done);
  const done = visible.filter((x) => x.done).sort((a, b) => (b.doneAt ?? 0) - (a.doneAt ?? 0));
  const groups = new Map<SectionKey, Task[]>(SECTIONS.map((k) => [k, []]));
  for (const task of open) groups.get(sectionOf(task, now))!.push(task);
  for (const [key, arr] of groups) {
    arr.sort((a, b) => (key === 'overdue' || key === 'upcoming' ? (a.date ?? '').localeCompare(b.date ?? '') : 0) || a.order - b.order);
  }
  const listById = new Map(s.lists.map((l) => [l.id, l]));
  const currentList = listId ? listById.get(listId) : undefined;

  const sectionDate = (key: SectionKey) => (key === 'today' ? now : key === 'tomorrow' ? addDays(now, 1) : key === 'upcoming' ? addDays(now, 2) : undefined);

  function onDragStart(task: Task, e: RPointerEvent<HTMLElement>) {
    e.preventDefault();
    e.stopPropagation();
    e.currentTarget.setPointerCapture(e.pointerId);
    startY.current = e.clientY;
    setDrag({ id: task.id, dy: 0, target: null, before: true, section: null });
  }

  function onDragMove(e: RPointerEvent<HTMLElement>) {
    if (!drag) return;
    let target: string | null = null;
    let before = true;
    let section: SectionKey | null = null;
    for (const el of document.elementsFromPoint(e.clientX, e.clientY)) {
      const row = el.closest<HTMLElement>('[data-task-id]');
      if (row && row.dataset.taskId === drag.id) continue;
      if (row) {
        const r = row.getBoundingClientRect();
        target = row.dataset.taskId!;
        before = e.clientY < r.top + r.height / 2;
        section = row.dataset.section as SectionKey;
        break;
      }
      const sec = el.closest<HTMLElement>('[data-section]');
      if (sec) {
        section = sec.dataset.section as SectionKey;
        break;
      }
    }
    setDrag({ ...drag, dy: e.clientY - startY.current, target, before, section });
  }

  function onDragEnd() {
    if (drag?.section) moveTask(drag.id, drag.target, drag.before, drag.section);
    setDrag(null);
  }

  function share() {
    const name = currentList?.name ?? t('tasks');
    const lines = [...open, ...done].map((x) => `${x.done ? '☑' : '☐'} ${x.title}`);
    void shareText(name, [name, ...lines].join('\n'), t('copied'));
  }

  return (
    <div className="view">
      <header className="top">
        <div>
          <p className="eyebrow">{t('appName')}</p>
          <h1>{fmtDate(new Date(), { month: 'long', day: 'numeric', weekday: 'long' })}</h1>
        </div>
        <button className="icon-btn" onClick={share} aria-label={t('share')}>
          <Icon name="share" />
        </button>
      </header>

      <div className="chips" role="tablist">
        <button className={`chip ${!listId ? 'on' : ''}`} onClick={() => setListId(null)}>
          {t('all')} <b>{s.tasks.filter((x) => !x.done).length}</b>
        </button>
        {s.lists.map((l) => (
          <button key={l.id} className={`chip ${listId === l.id ? 'on' : ''}`} onClick={() => setListId(l.id)}>
            <i className="dot" style={{ background: l.color }} />
            {l.name} <b>{s.tasks.filter((x) => !x.done && x.listId === l.id).length}</b>
          </button>
        ))}
      </div>

      {open.length === 0 && done.length === 0 && <p className="empty">{t('emptyTasks')}</p>}

      {SECTIONS.map((key) => {
        const arr = groups.get(key)!;
        if (key === 'overdue' && arr.length === 0) return null;
        const dropHere = drag?.section === key && !drag.target;
        return (
          <section key={key} data-section={key} className={`section ${dropHere ? 'drop' : ''}`}>
            <div className="section-head">
              <h2 className={key === 'overdue' ? 'danger' : ''}>
                {t(key)} {arr.length > 0 && <span>{arr.length}</span>}
              </h2>
              {key !== 'overdue' && (
                <button className="icon-btn sm" onClick={() => onAdd(sectionDate(key), listId ?? undefined)} aria-label={`${t(key)} · ${t('addTask')}`}>
                  <Icon name="plus" size={18} />
                </button>
              )}
            </div>
            {arr.map((task) => (
              <TaskRow
                key={task.id}
                task={task}
                section={key}
                list={listById.get(task.listId)}
                showDate={key === 'overdue' || key === 'upcoming'}
                onEdit={onEdit}
                dragging={drag?.id === task.id ? drag.dy : null}
                dropMark={drag?.target === task.id ? (drag.before ? 'before' : 'after') : null}
                onDragStart={onDragStart}
                onDragMove={onDragMove}
                onDragEnd={onDragEnd}
              />
            ))}
          </section>
        );
      })}

      {done.length > 0 && (
        <section className="section">
          <div className="section-head">
            <button className="link" onClick={() => setShowDone(!showDone)}>
              <h2>
                {t('completed')} <span>{done.length}</span>
              </h2>
              <span className={`caret ${showDone ? 'open' : ''}`}>
                <Icon name="down" size={16} />
              </span>
            </button>
            {showDone && (
              <button className="text-btn" onClick={() => clearCompleted(listId ?? undefined)}>
                {t('clearCompleted')}
              </button>
            )}
          </div>
          {showDone && done.map((task) => <TaskRow key={task.id} task={task} section="someday" list={listById.get(task.listId)} showDate={false} onEdit={onEdit} dragging={null} dropMark={null} />)}
        </section>
      )}
    </div>
  );
}

interface RowProps {
  task: Task;
  section: SectionKey;
  list?: TaskList;
  showDate: boolean;
  onEdit: (task: Task) => void;
  dragging: number | null;
  dropMark: 'before' | 'after' | null;
  onDragStart?: (task: Task, e: RPointerEvent<HTMLElement>) => void;
  onDragMove?: (e: RPointerEvent<HTMLElement>) => void;
  onDragEnd?: () => void;
}

function TaskRow({ task, section, list, showDate, onEdit, dragging, dropMark, onDragStart, onDragMove, onDragEnd }: RowProps) {
  const [dx, setDx] = useState(0);
  const swipe = useRef<{ x: number; y: number; active: boolean; moved: boolean } | null>(null);
  const subDone = task.subtasks.filter((x) => x.done).length;

  function down(e: RPointerEvent<HTMLDivElement>) {
    if (task.done || (e.target as HTMLElement).closest('button')) return;
    swipe.current = { x: e.clientX, y: e.clientY, active: false, moved: false };
  }
  function move(e: RPointerEvent<HTMLDivElement>) {
    const sw = swipe.current;
    if (!sw) return;
    const ddx = e.clientX - sw.x;
    const ddy = e.clientY - sw.y;
    if (!sw.active) {
      if (Math.abs(ddx) < 12 || Math.abs(ddx) < Math.abs(ddy)) return;
      sw.active = true;
      sw.moved = true;
      e.currentTarget.setPointerCapture(e.pointerId);
    }
    setDx(ddx);
  }
  function up() {
    const sw = swipe.current;
    if (sw?.active) {
      if (dx > 90) completeTask(task);
      else if (dx < -90) removeTask(task);
    }
    setDx(0);
    // 클릭 이벤트가 지나간 뒤에 지워야 밀기 직후 편집 창이 열리지 않는다.
    setTimeout(() => (swipe.current = null), 0);
  }

  const meta: string[] = [];
  if (showDate && task.date) meta.push(fmtDate(parseYmd(task.date), { month: 'short', day: 'numeric', weekday: 'short' }));
  if (task.time) meta.push(fmtTime(task.time));
  if (task.subtasks.length) meta.push(`${subDone}/${task.subtasks.length}`);

  return (
    <div
      className={`row-wrap ${dragging !== null ? 'dragging' : ''} ${dropMark ? `drop-${dropMark}` : ''}`}
      data-task-id={task.done ? undefined : task.id}
      data-section={section}
      style={dragging !== null ? { transform: `translateY(${dragging}px)` } : undefined}
    >
      {dx !== 0 && (
        <div className={`swipe-bg ${dx > 0 ? 'ok' : 'del'}`}>
          <Icon name={dx > 0 ? 'check' : 'trash'} />
        </div>
      )}
      <div
        className={`row ${task.done ? 'is-done' : ''}`}
        style={dx ? { transform: `translateX(${dx}px)`, transition: 'none' } : undefined}
        onPointerDown={down}
        onPointerMove={move}
        onPointerUp={up}
        onPointerCancel={up}
        onClick={() => {
          if (!swipe.current?.moved) onEdit(task);
        }}
      >
        <button
          className={`check p${task.priority} ${task.done ? 'on' : ''}`}
          role="checkbox"
          aria-checked={task.done}
          aria-label={task.title}
          onClick={(e) => {
            e.stopPropagation();
            completeTask(task);
          }}
        >
          {task.done && <Icon name="check" size={14} />}
        </button>
        <div className="row-main">
          <p className="row-title">{task.title}</p>
          {(meta.length > 0 || task.repeat !== 'none' || task.remind !== null || list) && (
            <p className="row-meta">
              {list && <i className="dot" style={{ background: list.color }} />}
              {meta.join(' · ')}
              {task.repeat !== 'none' && <Icon name="repeat" size={13} />}
              {task.remind !== null && <Icon name="bell" size={13} />}
            </p>
          )}
        </div>
        {onDragStart && !task.done && (
          <button className="grip" aria-label="reorder" onPointerDown={(e) => onDragStart(task, e)} onPointerMove={onDragMove} onPointerUp={onDragEnd} onPointerCancel={onDragEnd} onClick={(e) => e.stopPropagation()}>
            ⋮⋮
          </button>
        )}
      </div>
    </div>
  );
}
