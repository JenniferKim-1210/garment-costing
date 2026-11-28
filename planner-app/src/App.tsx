import { useEffect, useState } from 'react';
import { CalendarView } from './CalendarView';
import { today } from './date';
import { FocusView, useFocusTimer } from './FocusView';
import { setLang, t } from './i18n';
import { ItemSheet, newEvent } from './ItemSheet';
import type { Draft } from './ItemSheet';
import { startReminders } from './notify';
import { SettingsView } from './SettingsView';
import { getState, newTask, subscribe, useStore } from './store';
import { TasksView } from './TasksView';
import { Icon, toast, ToastHost } from './ui';
import type { IconName } from './ui';

type Tab = 'tasks' | 'calendar' | 'focus' | 'settings';
const TABS: Tab[] = ['tasks', 'calendar', 'focus', 'settings'];
const TAB_ICON: Record<Tab, IconName> = { tasks: 'tasks', calendar: 'calendar', focus: 'timer', settings: 'settings' };

export function App() {
  const s = useStore();
  const [tab, setTab] = useState<Tab>('tasks');
  const [draft, setDraft] = useState<Draft | null>(null);
  const timer = useFocusTimer();
  setLang(s.settings.lang);

  useEffect(() => {
    document.documentElement.dataset.theme = s.settings.theme;
  }, [s.settings.theme]);

  useEffect(() => startReminders(getState, subscribe, (r) => toast(`🔔 ${r.title}`)), []);

  const fab = () =>
    setDraft(tab === 'calendar' ? { kind: 'event', isNew: true, event: newEvent(today()) } : { kind: 'task', isNew: true, task: newTask({ date: today() }) });

  return (
    <div className="app">
      <main>
        {tab === 'tasks' && (
          <TasksView
            onEdit={(task) => setDraft({ kind: 'task', isNew: false, task })}
            onAdd={(date, listId) => setDraft({ kind: 'task', isNew: true, task: newTask({ date, ...(listId ? { listId } : {}) }) })}
          />
        )}
        {tab === 'calendar' && (
          <CalendarView
            onEditTask={(task) => setDraft({ kind: 'task', isNew: false, task })}
            onEditEvent={(event) => setDraft({ kind: 'event', isNew: false, event })}
            onAddEvent={(date, start) => setDraft({ kind: 'event', isNew: true, event: newEvent(date, start) })}
          />
        )}
        {tab === 'focus' && <FocusView timer={timer} />}
        {tab === 'settings' && <SettingsView />}
      </main>

      {(tab === 'tasks' || tab === 'calendar') && (
        <button className="fab" onClick={fab} aria-label={tab === 'calendar' ? t('addEvent') : t('addTask')}>
          <Icon name="plus" size={26} />
        </button>
      )}

      <nav className="tabbar">
        {TABS.map((k) => (
          <button key={k} className={tab === k ? 'on' : ''} onClick={() => setTab(k)} aria-current={tab === k ? 'page' : undefined}>
            <Icon name={TAB_ICON[k]} size={22} />
            <span>{t(k)}</span>
            {k === 'focus' && timer.running && <i className="live-dot" />}
          </button>
        ))}
      </nav>

      {draft && <ItemSheet key={draft.kind === 'task' ? draft.task.id : draft.event.id} draft={draft} onClose={() => setDraft(null)} />}
      <ToastHost />
    </div>
  );
}
