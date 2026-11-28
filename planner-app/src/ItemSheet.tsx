import { useEffect, useRef, useState } from 'react';
import { fromMinutes, parseYmd, toMinutes, today } from './date';
import { fmtDate, fmtTime, locale, reminderLabel, REMIND_OPTIONS, t } from './i18n';
import { parseQuick } from './parse';
import { deleteEvent, newTask, saveEvent, saveTask, uid, useStore } from './store';
import { removeTask } from './TasksView';
import { COLORS } from './types';
import type { CalEvent, Repeat, Task } from './types';
import { Icon, shareText, toast } from './ui';

export type Draft = { kind: 'task'; isNew: boolean; task: Task } | { kind: 'event'; isNew: boolean; event: CalEvent };

const REPEATS: Repeat[] = ['none', 'daily', 'weekdays', 'weekly', 'monthly', 'yearly'];

export function newEvent(date: string, start?: string): CalEvent {
  const now = new Date();
  const s = start ?? fromMinutes(Math.min(23, now.getHours() + 1) * 60);
  return { id: uid(), title: '', note: '', location: '', date, allDay: false, start: s, end: fromMinutes(Math.min(toMinutes(s) + 60, 23 * 60 + 59)), color: COLORS[1], repeat: 'none', remind: 10 };
}

// 브라우저마다 이름이 다른 음성 인식 API
type Recognizer = { lang: string; interimResults: boolean; onresult: ((e: { results: { 0: { transcript: string } }[] }) => void) | null; onend: (() => void) | null; onerror: (() => void) | null; start(): void; stop(): void };
function getRecognizer(): (new () => Recognizer) | undefined {
  const w = window as unknown as { SpeechRecognition?: new () => Recognizer; webkitSpeechRecognition?: new () => Recognizer };
  return w.SpeechRecognition ?? w.webkitSpeechRecognition;
}

export function ItemSheet({ draft, onClose }: { draft: Draft; onClose: () => void }) {
  const s = useStore();
  const [kind, setKind] = useState(draft.kind);
  const [task, setTask] = useState<Task>(draft.kind === 'task' ? draft.task : newTask({ date: draft.event.date }));
  const [ev, setEv] = useState<CalEvent>(draft.kind === 'event' ? draft.event : newEvent(draft.task.date ?? today(), draft.task.time));
  const [useParsed, setUseParsed] = useState(true);
  const [listening, setListening] = useState(false);
  const [sub, setSub] = useState('');
  const titleRef = useRef<HTMLInputElement>(null);
  const rec = useRef<Recognizer | null>(null);
  const title = kind === 'task' ? task.title : ev.title;
  const setTitle = (v: string) => (kind === 'task' ? setTask({ ...task, title: v }) : setEv({ ...ev, title: v }));

  useEffect(() => {
    if (draft.isNew) titleRef.current?.focus();
    return () => rec.current?.stop();
  }, [draft.isNew]);

  // 새 항목은 제목에서 날짜·시간을 읽어 제안한다.
  const parsed = draft.isNew && useParsed ? parseQuick(title) : null;
  const hint = parsed && (parsed.date || parsed.time) && parsed.title !== title ? parsed : null;

  function save() {
    if (!title.trim()) return;
    if (kind === 'task') {
      const out = { ...task, title: task.title.trim() };
      if (hint) Object.assign(out, { title: hint.title, date: hint.date ?? out.date, time: hint.time ?? out.time });
      if (!out.date) out.time = undefined;
      if (!out.date) out.remind = null;
      saveTask(out);
    } else {
      const out = { ...ev, title: ev.title.trim() };
      if (hint) {
        out.title = hint.title;
        if (hint.date) out.date = hint.date;
        if (hint.time) {
          const len = Math.max(30, toMinutes(out.end) - toMinutes(out.start));
          out.start = hint.time;
          out.end = fromMinutes(toMinutes(hint.time) + len);
          out.allDay = false;
        }
      }
      if (toMinutes(out.end) <= toMinutes(out.start)) out.end = fromMinutes(toMinutes(out.start) + 30);
      saveEvent(out);
    }
    onClose();
  }

  function remove() {
    if (kind === 'task') removeTask(task);
    else {
      deleteEvent(ev.id);
      toast(t('deleted'), t('undo'), () => saveEvent(ev));
    }
    onClose();
  }

  function voice() {
    const R = getRecognizer();
    if (!R) return toast(t('voiceUnsupported'));
    if (listening) return rec.current?.stop();
    const r = new R();
    r.lang = locale();
    r.interimResults = false;
    r.onresult = (e) => setTitle(e.results[0][0].transcript);
    r.onend = () => setListening(false);
    r.onerror = () => setListening(false);
    rec.current = r;
    setListening(true);
    r.start();
  }

  function share() {
    const lines = kind === 'task'
      ? [task.title, task.date ? `${fmtDate(parseYmd(task.date), { dateStyle: 'full' })}${task.time ? ` ${fmtTime(task.time)}` : ''}` : '', ...task.subtasks.map((x) => `${x.done ? '☑' : '☐'} ${x.title}`), task.note]
      : [ev.title, `${fmtDate(parseYmd(ev.date), { dateStyle: 'full' })}${ev.allDay ? '' : ` ${fmtTime(ev.start)} – ${fmtTime(ev.end)}`}`, ev.location, ev.note];
    void shareText(title, lines.filter(Boolean).join('\n'), t('copied'));
  }

  function addSub() {
    if (!sub.trim()) return;
    setTask({ ...task, subtasks: [...task.subtasks, { id: uid(), title: sub.trim(), done: false }] });
    setSub('');
  }

  const remindSelect = (value: number | null, onChange: (v: number | null) => void) => (
    <select value={value === null ? '' : String(value)} onChange={(e) => onChange(e.target.value === '' ? null : Number(e.target.value))}>
      {REMIND_OPTIONS.map((m) => (
        <option key={String(m)} value={m === null ? '' : String(m)}>
          {reminderLabel(m)}
        </option>
      ))}
    </select>
  );
  const repeatSelect = (value: Repeat, onChange: (v: Repeat) => void) => (
    <select value={value} onChange={(e) => onChange(e.target.value as Repeat)}>
      {REPEATS.map((r) => (
        <option key={r} value={r}>
          {t(`r_${r}`)}
        </option>
      ))}
    </select>
  );

  return (
    <div className="scrim" onClick={onClose}>
      <form
        className="sheet"
        role="dialog"
        aria-modal="true"
        onClick={(e) => e.stopPropagation()}
        onSubmit={(e) => {
          e.preventDefault();
          save();
        }}
      >
        <div className="sheet-bar">
          <button type="button" className="icon-btn" onClick={onClose} aria-label={t('cancel')}>
            <Icon name="x" />
          </button>
          {draft.isNew ? (
            <div className="seg">
              <button type="button" className={kind === 'task' ? 'on' : ''} onClick={() => { setTask({ ...task, title }); setKind('task'); }}>
                {t('task')}
              </button>
              <button type="button" className={kind === 'event' ? 'on' : ''} onClick={() => { setEv({ ...ev, title }); setKind('event'); }}>
                {t('event')}
              </button>
            </div>
          ) : (
            <span />
          )}
          <button type="submit" className="btn primary" disabled={!title.trim()}>
            {t('save')}
          </button>
        </div>

        <div className="sheet-body">
          <div className="title-row">
            <input ref={titleRef} className="title-input" value={title} onChange={(e) => setTitle(e.target.value)} placeholder={draft.isNew ? t('quickHint') : t('title')} aria-label={t('title')} enterKeyHint="done" />
            {draft.isNew && (
              <button type="button" className={`icon-btn ${listening ? 'live' : ''}`} onClick={voice} aria-label={t('voice')}>
                <Icon name="mic" />
              </button>
            )}
          </div>
          {listening && <p className="hint">{t('listening')}</p>}
          {hint && (
            <p className="hint">
              <Icon name="calendar" size={14} />
              {hint.date && fmtDate(parseYmd(hint.date), { month: 'short', day: 'numeric', weekday: 'short' })} {hint.time && fmtTime(hint.time)}
              <button type="button" className="icon-btn sm" onClick={() => setUseParsed(false)} aria-label={t('cancel')}>
                <Icon name="x" size={14} />
              </button>
            </p>
          )}

          {kind === 'task' ? (
            <>
              <label className="field">
                <span>{t('list')}</span>
                <select value={task.listId} onChange={(e) => setTask({ ...task, listId: e.target.value })}>
                  {s.lists.map((l) => (
                    <option key={l.id} value={l.id}>
                      {l.name}
                    </option>
                  ))}
                </select>
              </label>
              <div className="field-pair">
                <label className="field">
                  <span>{t('date')}</span>
                  <input type="date" value={task.date ?? ''} onChange={(e) => setTask({ ...task, date: e.target.value || undefined })} />
                </label>
                <label className="field">
                  <span>{t('time')}</span>
                  <input type="time" value={task.time ?? ''} disabled={!task.date} onChange={(e) => setTask({ ...task, time: e.target.value || undefined })} />
                </label>
              </div>
              <div className="field-pair">
                <label className="field">
                  <span>{t('reminder')}</span>
                  {remindSelect(task.remind, (v) => setTask({ ...task, remind: v, date: task.date ?? (v !== null ? today() : undefined) }))}
                </label>
                <label className="field">
                  <span>{t('repeat')}</span>
                  {repeatSelect(task.repeat, (v) => setTask({ ...task, repeat: v, date: task.date ?? (v !== 'none' ? today() : undefined) }))}
                </label>
              </div>
              <div className="field">
                <span>{t('priority')}</span>
                <div className="seg wide">
                  {([0, 1, 2, 3] as const).map((p) => (
                    <button type="button" key={p} className={task.priority === p ? 'on' : ''} onClick={() => setTask({ ...task, priority: p })}>
                      {t(`p${p}`)}
                    </button>
                  ))}
                </div>
              </div>
              <div className="field">
                <span>{t('subtasks')}</span>
                {task.subtasks.map((x) => (
                  <div key={x.id} className="sub">
                    <button type="button" className={`check ${x.done ? 'on' : ''}`} role="checkbox" aria-checked={x.done} aria-label={x.title} onClick={() => setTask({ ...task, subtasks: task.subtasks.map((y) => (y.id === x.id ? { ...y, done: !y.done } : y)) })}>
                      {x.done && <Icon name="check" size={14} />}
                    </button>
                    <input value={x.title} aria-label={t('subtasks')} onChange={(e) => setTask({ ...task, subtasks: task.subtasks.map((y) => (y.id === x.id ? { ...y, title: e.target.value } : y)) })} />
                    <button type="button" className="icon-btn sm" aria-label={t('delete')} onClick={() => setTask({ ...task, subtasks: task.subtasks.filter((y) => y.id !== x.id) })}>
                      <Icon name="x" size={16} />
                    </button>
                  </div>
                ))}
                <div className="sub">
                  <span className="check ghost" />
                  <input
                    value={sub}
                    placeholder={t('addSubtask')}
                    aria-label={t('addSubtask')}
                    onChange={(e) => setSub(e.target.value)}
                    onBlur={addSub}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        e.preventDefault();
                        addSub();
                      }
                    }}
                  />
                </div>
              </div>
              <label className="field">
                <span>{t('note')}</span>
                <textarea rows={3} value={task.note} onChange={(e) => setTask({ ...task, note: e.target.value })} />
              </label>
            </>
          ) : (
            <>
              <div className="field-pair">
                <label className="field">
                  <span>{t('date')}</span>
                  <input type="date" required value={ev.date} onChange={(e) => e.target.value && setEv({ ...ev, date: e.target.value })} />
                </label>
                <label className="field check-field">
                  <input type="checkbox" checked={ev.allDay} onChange={(e) => setEv({ ...ev, allDay: e.target.checked })} />
                  <span>{t('allDay')}</span>
                </label>
              </div>
              {!ev.allDay && (
                <div className="field-pair">
                  <label className="field">
                    <span>{t('start')}</span>
                    <input
                      type="time"
                      value={ev.start}
                      onChange={(e) => {
                        if (!e.target.value) return;
                        // 시작을 옮기면 길이를 유지한 채 종료도 따라 옮긴다.
                        const len = toMinutes(ev.end) - toMinutes(ev.start);
                        setEv({ ...ev, start: e.target.value, end: fromMinutes(toMinutes(e.target.value) + Math.max(len, 15)) });
                      }}
                    />
                  </label>
                  <label className="field">
                    <span>{t('end')}</span>
                    <input type="time" value={ev.end} onChange={(e) => e.target.value && setEv({ ...ev, end: e.target.value })} />
                  </label>
                </div>
              )}
              <div className="field-pair">
                <label className="field">
                  <span>{t('reminder')}</span>
                  {remindSelect(ev.remind, (v) => setEv({ ...ev, remind: v }))}
                </label>
                <label className="field">
                  <span>{t('repeat')}</span>
                  {repeatSelect(ev.repeat, (v) => setEv({ ...ev, repeat: v }))}
                </label>
              </div>
              <div className="field">
                <span>{t('color')}</span>
                <div className="swatches">
                  {COLORS.map((c) => (
                    <button type="button" key={c} className={ev.color === c ? 'on' : ''} style={{ background: c }} aria-label={c} aria-pressed={ev.color === c} onClick={() => setEv({ ...ev, color: c })} />
                  ))}
                </div>
              </div>
              <label className="field">
                <span>{t('location')}</span>
                <input value={ev.location} onChange={(e) => setEv({ ...ev, location: e.target.value })} />
              </label>
              <label className="field">
                <span>{t('note')}</span>
                <textarea rows={3} value={ev.note} onChange={(e) => setEv({ ...ev, note: e.target.value })} />
              </label>
            </>
          )}

          {!draft.isNew && (
            <div className="sheet-actions">
              <button type="button" className="btn" onClick={share}>
                <Icon name="share" size={18} /> {t('share')}
              </button>
              <button type="button" className="btn danger" onClick={remove}>
                <Icon name="trash" size={18} /> {t('delete')}
              </button>
            </div>
          )}
        </div>
      </form>
    </div>
  );
}
