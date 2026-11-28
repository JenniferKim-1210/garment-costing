import { useEffect, useRef, useState } from 'react';
import { today } from './date';
import { t } from './i18n';
import { showNow } from './notify';
import { getState, setState, useStore } from './store';
import { completeTask } from './TasksView';
import { toast } from './ui';

export interface FocusTimer {
  mode: 'focus' | 'break';
  running: boolean;
  remaining: number;
  total: number;
  taskId: string;
  setTaskId: (id: string) => void;
  toggle: () => void;
  reset: () => void;
}

/** 탭을 옮겨도 타이머가 유지되도록 App에서 한 번만 호출한다. */
export function useFocusTimer(): FocusTimer {
  const { settings } = useStore();
  const [mode, setMode] = useState<'focus' | 'break'>('focus');
  const [endAt, setEndAt] = useState<number | null>(null);
  const [paused, setPaused] = useState<number | null>(null);
  const [taskId, setTaskId] = useState('');
  const [, force] = useState(0);
  const total = (mode === 'focus' ? settings.focusMin : settings.breakMin) * 60000;
  const remaining = endAt ? Math.max(0, endAt - Date.now()) : paused ?? total;
  const modeRef = useRef(mode);
  modeRef.current = mode;

  useEffect(() => {
    if (!endAt) return;
    // 종료 시각 기준으로 계산하므로 화면이 꺼졌다 켜져도 시간이 어긋나지 않는다.
    const id = setInterval(() => {
      if (Date.now() < endAt) return force((n) => n + 1);
      const finished = modeRef.current;
      setEndAt(null);
      setPaused(null);
      setMode(finished === 'focus' ? 'break' : 'focus');
      const msg = finished === 'focus' ? t('focusDone') : t('breakDone');
      if (finished === 'focus') {
        const day = today();
        setState((s) => ({ ...s, focusLog: { ...s.focusLog, [day]: (s.focusLog[day] ?? 0) + 1 } }));
      }
      toast(msg);
      void showNow(t('focusTitle'), msg);
      navigator.vibrate?.(300);
    }, 250);
    return () => clearInterval(id);
  }, [endAt]);

  return {
    mode,
    running: endAt !== null,
    remaining,
    total,
    taskId,
    setTaskId,
    toggle: () => {
      if (endAt) {
        setPaused(Math.max(0, endAt - Date.now()));
        setEndAt(null);
      } else {
        setEndAt(Date.now() + (paused ?? total));
        setPaused(null);
      }
    },
    reset: () => {
      setEndAt(null);
      setPaused(null);
      setMode('focus');
    },
  };
}

export function FocusView({ timer }: { timer: FocusTimer }) {
  const s = useStore();
  const now = today();
  const candidates = s.tasks.filter((x) => !x.done && (!x.date || x.date <= now));
  const task = s.tasks.find((x) => x.id === timer.taskId && !x.done);
  const sec = Math.ceil(timer.remaining / 1000);
  const R = 120;
  const C = 2 * Math.PI * R;
  const progress = timer.total ? 1 - timer.remaining / timer.total : 0;

  return (
    <div className="view focus">
      <header className="top">
        <div>
          <p className="eyebrow">{t('focusTitle')}</p>
          <h1>{timer.mode === 'focus' ? t('focusing') : t('break')}</h1>
        </div>
      </header>

      <label className="field">
        <span>{t('focusOn')}</span>
        <select value={task ? timer.taskId : ''} onChange={(e) => timer.setTaskId(e.target.value)}>
          <option value="">{t('focusNone')}</option>
          {candidates.map((x) => (
            <option key={x.id} value={x.id}>
              {x.title}
            </option>
          ))}
        </select>
      </label>

      <div className={`dial ${timer.mode}`} role="timer" aria-live="off">
        <svg viewBox="0 0 280 280">
          <circle cx="140" cy="140" r={R} className="track" />
          <circle cx="140" cy="140" r={R} className="bar" strokeDasharray={C} strokeDashoffset={C * (1 - progress)} transform="rotate(-90 140 140)" />
        </svg>
        <div className="dial-text">
          <strong>
            {String(Math.floor(sec / 60)).padStart(2, '0')}:{String(sec % 60).padStart(2, '0')}
          </strong>
          {task && <span>{task.title}</span>}
        </div>
      </div>

      <div className="focus-btns">
        <button className="btn" onClick={timer.reset}>
          {t('reset')}
        </button>
        <button className="btn primary big" onClick={timer.toggle}>
          {timer.running ? t('pause') : t('startBtn')}
        </button>
        {task && (
          <button
            className="btn"
            onClick={() => {
              completeTask(getState().tasks.find((x) => x.id === task.id)!);
              timer.setTaskId('');
            }}
          >
            {t('markDone')}
          </button>
        )}
      </div>
      <p className="stat">
        {t('sessionsToday')} <b>{s.focusLog[now] ?? 0}</b>
      </p>
    </div>
  );
}
