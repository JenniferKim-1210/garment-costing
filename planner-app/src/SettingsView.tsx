import { useEffect, useRef, useState } from 'react';
import { t } from './i18n';
import { getPermission, isNative, requestPermission } from './notify';
import type { Permission } from './notify';
import { deleteList, getState, normalize, saveList, setState, uid, useStore, wipe } from './store';
import { COLORS } from './types';
import type { Settings } from './types';
import { Icon, toast } from './ui';

export function SettingsView() {
  const s = useStore();
  const [perm, setPerm] = useState<Permission>('default');
  const [name, setName] = useState('');
  const file = useRef<HTMLInputElement>(null);
  const set = (patch: Partial<Settings>) => setState((st) => ({ ...st, settings: { ...st.settings, ...patch } }));

  useEffect(() => {
    void getPermission().then(setPerm);
  }, []);

  function exportData() {
    const blob = new Blob([JSON.stringify(getState(), null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `haru-planner-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }

  async function importData(f: File | undefined) {
    if (!f) return;
    try {
      const next = normalize(JSON.parse(await f.text()));
      if (!next) throw new Error('bad format');
      setState(() => next);
      toast(t('importDone'));
    } catch {
      toast(t('importFail'));
    }
  }

  function addList() {
    if (!name.trim()) return;
    saveList({ id: uid(), name: name.trim(), color: COLORS[s.lists.length % COLORS.length] });
    setName('');
  }

  const num = (v: string, min: number, max: number) => Math.min(max, Math.max(min, Math.round(Number(v)) || min));

  return (
    <div className="view">
      <header className="top">
        <div>
          <p className="eyebrow">{t('appName')}</p>
          <h1>{t('settings')}</h1>
        </div>
      </header>

      <section className="card">
        <label className="field inline">
          <span>{t('language')}</span>
          <select value={s.settings.lang} onChange={(e) => set({ lang: e.target.value as Settings['lang'] })}>
            <option value="ko">한국어</option>
            <option value="en">English</option>
          </select>
        </label>
        <label className="field inline">
          <span>{t('theme')}</span>
          <select value={s.settings.theme} onChange={(e) => set({ theme: e.target.value as Settings['theme'] })}>
            <option value="system">{t('themeSystem')}</option>
            <option value="light">{t('themeLight')}</option>
            <option value="dark">{t('themeDark')}</option>
          </select>
        </label>
        <label className="field inline">
          <span>{t('weekStart')}</span>
          <select value={s.settings.weekStart} onChange={(e) => set({ weekStart: Number(e.target.value) as 0 | 1 })}>
            <option value={0}>{t('sunday')}</option>
            <option value={1}>{t('monday')}</option>
          </select>
        </label>
        <label className="field inline">
          <span>{t('focusLen')}</span>
          <input type="number" inputMode="numeric" min={1} max={180} value={s.settings.focusMin} onChange={(e) => set({ focusMin: num(e.target.value, 1, 180) })} />
        </label>
        <label className="field inline">
          <span>{t('breakLen')}</span>
          <input type="number" inputMode="numeric" min={1} max={60} value={s.settings.breakMin} onChange={(e) => set({ breakMin: num(e.target.value, 1, 60) })} />
        </label>
      </section>

      <section className="card">
        <h2>{t('notifications')}</h2>
        {perm === 'granted' ? (
          <p className="note">{t('notifOn')}</p>
        ) : perm === 'denied' || perm === 'unsupported' ? (
          <p className="note">{t('notifDenied')}</p>
        ) : (
          <button className="btn primary" onClick={() => void requestPermission().then(setPerm)}>
            <Icon name="bell" size={18} /> {t('notifEnable')}
          </button>
        )}
        {!isNative && <p className="note">{t('notifWebNote')}</p>}
      </section>

      <section className="card">
        <h2>{t('lists')}</h2>
        {s.lists.map((l) => (
          <div key={l.id} className="list-row">
            <button
              className="swatch"
              style={{ background: l.color }}
              aria-label={t('color')}
              onClick={() => saveList({ ...l, color: COLORS[(COLORS.indexOf(l.color) + 1) % COLORS.length] })}
            />
            <input value={l.name} aria-label={t('listName')} onChange={(e) => saveList({ ...l, name: e.target.value })} />
            {s.lists.length > 1 && (
              <button className="icon-btn sm" aria-label={t('delete')} onClick={() => confirm(t('deleteListConfirm')) && deleteList(l.id)}>
                <Icon name="trash" size={16} />
              </button>
            )}
          </div>
        ))}
        <form
          className="list-row"
          onSubmit={(e) => {
            e.preventDefault();
            addList();
          }}
        >
          <span className="swatch ghost" />
          <input value={name} placeholder={t('newList')} aria-label={t('newList')} onChange={(e) => setName(e.target.value)} />
          <button className="icon-btn sm" aria-label={t('newList')} disabled={!name.trim()}>
            <Icon name="plus" size={18} />
          </button>
        </form>
      </section>

      <section className="card">
        <h2>{t('backup')}</h2>
        <p className="note">{t('dataNote')}</p>
        <div className="btn-row">
          <button className="btn" onClick={exportData}>
            {t('export')}
          </button>
          <button className="btn" onClick={() => file.current?.click()}>
            {t('import')}
          </button>
          <input ref={file} type="file" accept="application/json,.json" hidden onChange={(e) => { void importData(e.target.files?.[0]); e.target.value = ''; }} />
        </div>
        <button className="btn danger" onClick={() => confirm(t('wipeConfirm')) && wipe()}>
          {t('wipe')}
        </button>
      </section>

      <p className="foot">
        <a href="./privacy.html" target="_blank" rel="noreferrer">
          {t('privacy')}
        </a>{' '}
        · v1.0.0
      </p>
    </div>
  );
}
