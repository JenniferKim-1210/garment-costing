import { useEffect, useState } from 'react';

const PATHS = {
  check: 'M5 12l5 5 9-10',
  tasks: 'M9 6h11M9 12h11M9 18h11M4 6h.01M4 12h.01M4 18h.01',
  calendar: 'M7 3v3M17 3v3M4 9h16M5 5h14a1 1 0 011 1v13a1 1 0 01-1 1H5a1 1 0 01-1-1V6a1 1 0 011-1z',
  timer: 'M12 8v5l3 2M9 2h6M12 5a8 8 0 100 16 8 8 0 000-16z',
  settings: 'M4 6h10M18 6h2M4 12h4M12 12h8M4 18h12M14 4v4M8 10v4M16 16v4',
  plus: 'M12 5v14M5 12h14',
  mic: 'M12 3a3 3 0 00-3 3v6a3 3 0 006 0V6a3 3 0 00-3-3zM5 11a7 7 0 0014 0M12 18v3',
  left: 'M15 5l-7 7 7 7',
  right: 'M9 5l7 7-7 7',
  down: 'M5 9l7 7 7-7',
  repeat: 'M4 12a8 8 0 0114-5l2 2M20 4v5h-5M20 12a8 8 0 01-14 5l-2-2M4 20v-5h5',
  bell: 'M6 16v-5a6 6 0 1112 0v5l2 2H4l2-2zM10 21h4',
  trash: 'M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13',
  share: 'M12 15V4M8 8l4-4 4 4M5 13v6h14v-6',
  x: 'M6 6l12 12M18 6L6 18',
  pin: 'M12 21s7-6.2 7-11a7 7 0 10-14 0c0 4.800 7 11 7 11zM12 12.500a2.500 2.500 0 100-5 2.500 2.500 0 000 5z',
};

export type IconName = keyof typeof PATHS;

export function Icon({ name, size = 20 }: { name: IconName; size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d={PATHS[name]} />
    </svg>
  );
}

// ---------- 토스트 ----------

interface ToastData {
  id: number;
  msg: string;
  actionLabel?: string;
  action?: () => void;
}

let push: ((t: ToastData) => void) | null = null;

export function toast(msg: string, actionLabel?: string, action?: () => void) {
  push?.({ id: Date.now(), msg, actionLabel, action });
}

export function ToastHost() {
  const [item, setItem] = useState<ToastData | null>(null);
  useEffect(() => {
    push = setItem;
    return () => {
      push = null;
    };
  }, []);
  useEffect(() => {
    if (!item) return;
    const id = setTimeout(() => setItem(null), 4500);
    return () => clearTimeout(id);
  }, [item]);
  if (!item) return null;
  return (
    <div className="toast" role="status">
      <span>{item.msg}</span>
      {item.action && (
        <button
          onClick={() => {
            item.action?.();
            setItem(null);
          }}
        >
          {item.actionLabel}
        </button>
      )}
    </div>
  );
}

export async function shareText(title: string, text: string, copiedMsg: string) {
  try {
    if (navigator.share) {
      await navigator.share({ title, text });
      return;
    }
    await navigator.clipboard.writeText(text);
    toast(copiedMsg);
  } catch {
    // 사용자가 공유 창을 닫은 경우는 조용히 넘어간다.
  }
}
