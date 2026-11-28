/* 단어카드 앱 로직
 * - 저장: 휴대폰 안(localStorage)에만 저장합니다. 서버로 보내지 않습니다.
 * - 복습 간격: 알아요를 누를수록 1일 → 2일 → 4일 → 7일 → 15일 → 30일 … 로 길어집니다.
 * - 안드로이드 앱(Capacitor)으로 실행되면 알림·발음은 네이티브 기능을 씁니다.
 */
(function () {
  'use strict';

  // ---------- 기본 도구 ----------
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* 저장 불가 환경 */ } }
  };
  const DAY = 86400000;
  const dayNum = (date = new Date()) => Math.floor((date.getTime() - date.getTimezoneOffset() * 60000) / DAY);
  const today = () => dayNum();
  const dayToDate = (n) => { const d = new Date(n * DAY); return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()); };
  const shuffle = (a) => { a = a.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };

  const NATIVE = !!(window.Capacitor && window.Capacitor.isNativePlatform && window.Capacitor.isNativePlatform());
  const plugin = (name) => {
    if (!NATIVE) return null;
    const C = window.Capacitor;
    return (C.Plugins && C.Plugins[name]) || (C.registerPlugin ? C.registerPlugin(name) : null);
  };

  // 복습 간격(일)
  const INTERVALS = [0, 1, 2, 4, 7, 15, 30, 60, 120];
  const KNOWN_BOX = 4; // 7일 간격 이상이면 '외움'
  const POINTS = { attend: 10, streakStep: 2, streakMax: 20, card: 1, goal: 20, quiz: 2, week: 30, month: 100 };
  const LEVEL_SIZE = 200;

  // ---------- 상태 ----------
  const DEFAULT_SETTINGS = { deck: null, daily: 20, dir: 'en', autoSpeak: true, rate: 0.9, remind: true, remindHour: 20 };
  let settings = Object.assign({}, DEFAULT_SETTINGS, store.get('wc.settings', {}));
  let progress = store.get('wc.progress', {});
  let custom = store.get('wc.custom', []);
  let stats = Object.assign({ pts: 0, attend: [], studied: [], goal: [], newBy: {} }, store.get('wc.stats', {}));

  function save() {
    store.set('wc.settings', settings);
    store.set('wc.progress', progress);
    store.set('wc.custom', custom);
    store.set('wc.stats', stats);
    syncReminder();
  }

  // ---------- 단어장 ----------
  const BUILTIN = (window.WORDCARD_DECKS || []).map((d) => ({
    id: d.id, name: d.name, desc: d.desc,
    cards: d.words.trim().split('\n').map((line) => {
      const [en, ko, ex] = line.split('|').map((s) => (s || '').trim());
      return { id: d.id + ':' + en.toLowerCase(), en, ko, ex: ex || '', deck: d.id };
    })
  }));
  function decks() {
    return BUILTIN.concat([{
      id: 'mine', name: '내 단어장', desc: '직접 추가한 단어',
      cards: custom.map((c) => ({ id: c.id, en: c.en, ko: c.ko, ex: c.ex || '', deck: 'mine' }))
    }]);
  }
  const deckById = (id) => decks().find((d) => d.id === id) || decks()[0];
  const allCards = () => decks().flatMap((d) => d.cards);
  const cardsOf = (id) => (id === 'all' ? allCards() : deckById(id).cards);
  function status(c) {
    const p = progress[c.id];
    if (!p) return 'new';
    return p.b >= KNOWN_BOX ? 'known' : 'learning';
  }
  const STATUS_LABEL = { new: '새 단어', learning: '외우는 중', known: '외움' };

  // ---------- 날짜 기록 ----------
  const hasDay = (arr, d) => arr.includes(d);
  function addDay(arr, d) { if (!arr.includes(d)) { arr.push(d); arr.sort((a, b) => a - b); if (arr.length > 800) arr.splice(0, arr.length - 800); } }
  function streakOf(arr) {
    let d = today();
    if (!hasDay(arr, d)) d -= 1;
    let n = 0;
    while (hasDay(arr, d)) { n++; d--; }
    return n;
  }
  const lastDay = (arr) => (arr.length ? arr[arr.length - 1] : null);
  function pruneNewBy() {
    const t = today();
    Object.keys(stats.newBy).forEach((k) => { if (+k < t - 40) delete stats.newBy[k]; });
  }

  // ---------- 점수 ----------
  function addPoints(n) { stats.pts = Math.max(0, (stats.pts || 0) + n); renderPoints(); }
  function level() {
    const p = stats.pts || 0;
    return { lv: Math.floor(p / LEVEL_SIZE) + 1, cur: p % LEVEL_SIZE, size: LEVEL_SIZE };
  }
  function checkIn(auto) {
    const t = today();
    if (hasDay(stats.attend, t)) { if (!auto) toast('오늘은 이미 출석했어요'); return 0; }
    addDay(stats.attend, t);
    const s = streakOf(stats.attend);
    let gain = POINTS.attend + Math.min(POINTS.streakMax, (s - 1) * POINTS.streakStep);
    let msg = `출석 완료 +${gain}점`;
    if (s > 0 && s % 30 === 0) { gain += POINTS.month; msg += ` · ${s}일 연속 보너스 +${POINTS.month}`; }
    else if (s > 0 && s % 7 === 0) { gain += POINTS.week; msg += ` · ${s}일 연속 보너스 +${POINTS.week}`; }
    addPoints(gain);
    save();
    toast(msg);
    return gain;
  }

  // ---------- 간격 반복 ----------
  function newToday() { return stats.newBy[today()] || 0; }
  function dueCards(deckId) {
    const t = today();
    return cardsOf(deckId).filter((c) => progress[c.id] && progress[c.id].d <= t)
      .sort((a, b) => progress[a.id].d - progress[b.id].d);
  }
  function freshCards(deckId, n) { return cardsOf(deckId).filter((c) => !progress[c.id]).slice(0, Math.max(0, n)); }
  function buildQueue(deckId, extraNew) {
    const due = dueCards(deckId);
    const allowance = extraNew != null ? extraNew : settings.daily - newToday();
    const fresh = freshCards(deckId, allowance);
    // 새 단어 2개마다 복습 1개를 섞습니다
    const q = [];
    let i = 0, j = 0;
    while (i < fresh.length || j < due.length) {
      if (i < fresh.length) q.push(fresh[i++]);
      if (i < fresh.length) q.push(fresh[i++]);
      if (j < due.length) q.push(due[j++]);
    }
    return q;
  }
  function preview(c, g) {
    const p = progress[c.id] || { b: 0 };
    if (g === 0) return '다시';
    if (g === 1) { const b = Math.max(1, p.b); return Math.max(1, Math.round(INTERVALS[b] / 2)) + '일 뒤'; }
    const b = Math.min(INTERVALS.length - 1, p.b + 1); return INTERVALS[b] + '일 뒤';
  }
  function grade(c, g) {
    const t = today();
    let p = progress[c.id];
    if (!p) {
      p = { b: 0, d: t, r: 0, l: 0 };
      stats.newBy[t] = (stats.newBy[t] || 0) + 1;
    }
    p.r++;
    if (g === 0) { p.b = 0; p.l++; p.d = t; }
    else if (g === 1) { p.b = Math.max(1, p.b); p.d = t + Math.max(1, Math.round(INTERVALS[p.b] / 2)); }
    else { p.b = Math.min(INTERVALS.length - 1, p.b + 1); p.d = t + INTERVALS[p.b]; }
    progress[c.id] = p;
    addDay(stats.studied, t);
    if (g > 0) addPoints(POINTS.card);
    // 오늘 목표 달성
    if (newToday() >= settings.daily && !hasDay(stats.goal, t)) {
      addDay(stats.goal, t);
      addPoints(POINTS.goal);
      toast(`오늘 목표 ${settings.daily}개 달성! +${POINTS.goal}점`);
    }
    pruneNewBy();
    save();
  }

  // ---------- 발음 ----------
  let voices = [];
  function loadVoices() { try { voices = speechSynthesis.getVoices().filter((v) => /^en(-|_)/i.test(v.lang)); } catch (e) { voices = []; } }
  if ('speechSynthesis' in window) { loadVoices(); try { speechSynthesis.onvoiceschanged = loadVoices; } catch (e) { /* noop */ } }
  function speak(text) {
    if (!text) return;
    const tts = plugin('TextToSpeech');
    if (tts) { tts.stop().catch(() => {}).finally(() => tts.speak({ text, lang: 'en-US', rate: settings.rate, pitch: 1, volume: 1 }).catch(() => toast('발음을 재생할 수 없어요'))); return; }
    if (!('speechSynthesis' in window)) { toast('이 기기는 발음 듣기를 지원하지 않아요'); return; }
    try {
      speechSynthesis.cancel();
      const u = new SpeechSynthesisUtterance(text);
      u.lang = 'en-US'; u.rate = settings.rate;
      const v = voices.find((x) => /en-US/i.test(x.lang)) || voices[0];
      if (v) u.voice = v;
      speechSynthesis.speak(u);
    } catch (e) { toast('발음을 재생할 수 없어요'); }
  }

  // ---------- 알림 (공부 안 한 날) ----------
  const REMIND_IDS = [1001, 1002, 1003, 1004, 1005, 1006, 1007];
  function remindText() {
    const s = streakOf(stats.studied);
    return s > 0 ? `연속 ${s}일째 공부 중이에요. 오늘 단어 ${settings.daily}개로 기록을 이어가세요!` : `오늘 단어 ${settings.daily}개를 아직 안 했어요. 5분이면 충분해요!`;
  }
  let reminderTimer = null;
  async function syncReminder() {
    const t = today();
    const studiedToday = hasDay(stats.studied, t);
    // 1) 안드로이드 앱: 앞으로 7일 치 알림을 예약 (공부한 날은 건너뜀)
    const LN = plugin('LocalNotifications');
    if (LN) {
      try {
        await LN.cancel({ notifications: REMIND_IDS.map((id) => ({ id })) });
        if (!settings.remind) return;
        const now = new Date();
        const list = [];
        for (let i = 0; i < 7; i++) {
          if (i === 0 && studiedToday) continue;
          const at = dayToDate(t + i); at.setHours(settings.remindHour, 0, 0, 0);
          if (at <= now) continue;
          list.push({ id: REMIND_IDS[i], title: '단어카드', body: i === 0 ? remindText() : `오늘 단어 ${settings.daily}개, 지금 시작해 볼까요?`, schedule: { at, allowWhileIdle: true } });
        }
        if (list.length) await LN.schedule({ notifications: list });
      } catch (e) { /* 권한 없음 */ }
      return;
    }
    // 2) 웹/설치형 앱: 서비스워커가 읽을 수 있게 상태 저장 + 앱이 켜져 있으면 시간 맞춰 알림
    try {
      if ('caches' in window) {
        const c = await caches.open('wordcard-state');
        await c.put('./__state', new Response(JSON.stringify({ remind: settings.remind, hour: settings.remindHour, studied: lastDay(stats.studied), body: remindText() }), { headers: { 'Content-Type': 'application/json' } }));
      }
    } catch (e) { /* noop */ }
    clearTimeout(reminderTimer);
    if (settings.remind && !studiedToday && 'Notification' in window && Notification.permission === 'granted') {
      const at = new Date(); at.setHours(settings.remindHour, 0, 0, 0);
      const ms = at - new Date();
      if (ms > 0 && ms < DAY) reminderTimer = setTimeout(() => { if (!hasDay(stats.studied, today())) showWebNotice(); }, ms);
    }
  }
  async function showWebNotice() {
    try {
      const reg = 'serviceWorker' in navigator ? await navigator.serviceWorker.getRegistration() : null;
      if (reg) reg.showNotification('단어카드', { body: remindText(), icon: 'icons/icon-192.png', badge: 'icons/icon-192.png', tag: 'study-reminder' });
      else new Notification('단어카드', { body: remindText(), icon: 'icons/icon-192.png' });
    } catch (e) { /* noop */ }
  }
  async function requestNotify() {
    const LN = plugin('LocalNotifications');
    if (LN) {
      try {
        const r = await LN.requestPermissions();
        if (r.display !== 'granted') { toast('휴대폰 설정에서 알림을 허용해 주세요'); return false; }
        return true;
      } catch (e) { return false; }
    }
    if (!('Notification' in window)) { toast('이 브라우저는 알림을 지원하지 않아요. 앱 안 알림만 표시돼요'); return false; }
    try {
      const r = await Notification.requestPermission();
      if (r !== 'granted') { toast('알림이 차단되어 있어요. 브라우저 설정에서 허용해 주세요'); return false; }
      try {
        const reg = await navigator.serviceWorker?.ready;
        if (reg && 'periodicSync' in reg) await reg.periodicSync.register('study-reminder', { minInterval: 6 * 60 * 60 * 1000 });
      } catch (e) { /* 지원 안 함 */ }
      return true;
    } catch (e) { return false; }
  }

  // ---------- 화면 공통 ----------
  let toastTimer = null;
  function toast(msg) {
    const el = $('#toast');
    el.textContent = msg; el.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { el.hidden = true; }, 2400);
  }
  function renderPoints() { $('#lvText').textContent = `Lv.${level().lv} · ${(stats.pts || 0).toLocaleString('ko-KR')}점`; }
  const ICON = {
    speak: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M11 5 6 9H3v6h3l5 4z"/><path d="M15.5 8.5a5 5 0 0 1 0 7M18.5 5.5a9 9 0 0 1 0 13"/></svg>',
    close: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" aria-hidden="true"><path d="M5 5l14 14M19 5 5 19"/></svg>',
    trash: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13"/></svg>',
    down: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 9l7 7 7-7"/></svg>',
    // 채점 기호: 모름 ✕ · 애매 △ · 알아요 ○
    x: '<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="3.6" stroke-linecap="round" aria-hidden="true"><path d="M7 7l18 18M25 7 7 25"/></svg>',
    tri: '<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="3.6" stroke-linejoin="round" aria-hidden="true"><path d="M16 5.5 28 26H4z"/></svg>',
    o: '<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="3.6" aria-hidden="true"><circle cx="16" cy="16" r="11.5"/></svg>'
  };
  const CHECK = '<path d="M5 12.5l4.5 4.5L19 7.5" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/>';
  const checkSvg = (pop) => `<svg class="chk${pop ? ' pop' : ''}" viewBox="0 0 24 24" aria-hidden="true">${CHECK}</svg>`;
  const BIG_CHECK = `<div class="bigchk"><svg viewBox="0 0 24 24" aria-hidden="true">${CHECK}</svg></div>`;

  let currentView = 'home';
  function go(view) {
    currentView = view;
    ['home', 'quiz', 'words', 'settings'].forEach((v) => { $('#view-' + v).hidden = v !== view; });
    $$('.tab').forEach((b) => { if (b.dataset.view === view) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current'); });
    if (view === 'home') renderHome();
    if (view === 'quiz') renderQuizStart();
    if (view === 'words') renderWords();
    if (view === 'settings') renderSettings();
    window.scrollTo(0, 0);
  }
  $$('.tab').forEach((b) => b.addEventListener('click', () => go(b.dataset.view)));

  // 뒤로 가기 버튼으로 학습 화면 닫기 (안드로이드)
  let overlayOpen = null;
  function pushOverlay(name) { overlayOpen = name; try { history.pushState({ wc: name }, ''); } catch (e) { /* noop */ } }
  function popOverlay() { if (overlayOpen) { overlayOpen = null; try { history.back(); } catch (e) { /* noop */ } } }
  window.addEventListener('popstate', () => {
    if (overlayOpen === 'study') { overlayOpen = null; closeStudy(true); }
  });

  // ---------- 홈 ----------
  function weekStrip(pop) {
    const t = today();
    const names = ['일', '월', '화', '수', '목', '금', '토'];
    let html = '';
    for (let i = 6; i >= 0; i--) {
      const d = t - i;
      const date = dayToDate(d);
      const on = hasDay(stats.attend, d);
      html += `<div class="day${i === 0 ? ' is-today' : ''}"><span class="slot${on ? ' on' : ''}" aria-label="${date.getMonth() + 1}월 ${date.getDate()}일 ${on ? '출석' : '미출석'}">${on ? checkSvg(pop && i === 0) : date.getDate()}</span>${i === 0 ? '오늘' : names[date.getDay()]}</div>`;
    }
    return html;
  }
  function renderHome(popStamp) {
    const deck = deckById(settings.deck);
    const t = today();
    const done = Math.min(newToday(), settings.daily);
    const due = dueCards(deck.id).length;
    const fresh = freshCards(deck.id, 9999).length;
    const leftNew = Math.max(0, Math.min(settings.daily - newToday(), fresh));
    const streak = streakOf(stats.studied);
    const attStreak = streakOf(stats.attend);
    const attended = hasDay(stats.attend, t);
    const lv = level();
    const last = lastDay(stats.studied);

    let memo = '';
    if (last != null && last < t - 1) {
      memo = `<div class="memo" role="alert"><b>${t - last - 1}일 쉬었어요</b>${due ? `복습할 단어가 ${due}개 쌓였어요. ` : ''}오늘부터 다시 이어 가요.</div>`;
    } else if (last != null && !hasDay(stats.studied, t) && new Date().getHours() >= settings.remindHour) {
      memo = `<div class="memo" role="alert"><b>오늘 아직 안 했어요</b>${esc(remindText())}</div>`;
    }

    let headline, cta;
    if (leftNew > 0 && done === 0) {
      headline = `오늘의 단어<br><mark>${leftNew}개</mark> 준비됐어요`;
      cta = '<button class="cta" id="startStudy">시작하기</button>';
    } else if (leftNew > 0) {
      headline = `오늘 ${settings.daily}개 중<br><mark>${leftNew}개</mark> 남았어요`;
      cta = '<button class="cta" id="startStudy">이어서 하기</button>';
    } else if (due > 0) {
      headline = `새 단어는 끝!<br>복습 <mark>${due}개</mark> 남았어요`;
      cta = '<button class="cta" id="startStudy">복습하기</button>';
    } else if (fresh > 0) {
      headline = '오늘 공부<br><mark>다 했어요</mark>';
      cta = '<button class="line" id="moreStudy">10개 더 하기</button>';
    } else {
      headline = '이 단어장은<br><mark>끝까지 봤어요</mark>';
      cta = '<button class="line" id="otherDeck">다른 단어장 고르기</button>';
    }
    const cells = Array.from({ length: settings.daily }, (_, i) => `<i${i < done ? ' class="on"' : ''}></i>`).join('');
    const known = deck.cards.filter((c) => status(c) === 'known').length;
    const learning = deck.cards.filter((c) => status(c) === 'learning').length;

    $('#view-home').innerHTML = `
      <div class="hero">
        <button class="deckpick" id="switchDeck">${esc(deck.name)}${ICON.down}<span class="sr">단어장 바꾸기</span></button>
        ${memo}
        <h1 class="headline">${headline}</h1>
        <div class="tally" role="img" aria-label="오늘 새 단어 ${done}개 완료, 목표 ${settings.daily}개">${cells}</div>
        <p class="sub">${due ? `복습 ${due}개 · ` : ''}${streak > 0 ? `${streak}일째 이어서 공부 중` : '오늘 시작하면 1일째'}</p>
        ${cta}
      </div>
      <div class="band"></div>
      <div class="sec">
        <div class="sec-head"><h2>출석</h2><span>${attStreak > 0 ? `연속 ${attStreak}일` : '아직 없어요'}</span></div>
        <div class="week">${weekStrip(popStamp)}</div>
        ${attended
          ? '<button class="line done" id="checkIn" disabled>오늘 출석 완료</button>'
          : `<button class="line" id="checkIn">출석 체크 <b>+${POINTS.attend + Math.min(POINTS.streakMax, attStreak * POINTS.streakStep)}점</b></button>`}
        <div class="lvrow"><b>Lv.${lv.lv}</b><div class="lvbar"><i style="width:${Math.round((lv.cur / lv.size) * 100)}%"></i></div><span>다음 레벨까지 ${lv.size - lv.cur}점</span></div>
        <details class="how">
          <summary>점수는 이렇게 받아요</summary>
          <ul>
            <li>출석 <b>+${POINTS.attend}</b></li>
            <li>연속 출석 하루마다 <b>+${POINTS.streakStep} (최대 ${POINTS.streakMax})</b></li>
            <li>7일 연속 / 30일 연속 <b>+${POINTS.week} / +${POINTS.month}</b></li>
            <li>카드 1장 맞힘 <b>+${POINTS.card}</b></li>
            <li>퀴즈 1문제 정답 <b>+${POINTS.quiz}</b></li>
            <li>하루 ${settings.daily}개 달성 <b>+${POINTS.goal}</b></li>
          </ul>
        </details>
      </div>
      <div class="band"></div>
      <div class="sec">
        <div class="sec-head"><h2>지금까지</h2><span>${esc(deck.name)} ${deck.cards.length}개</span></div>
        <div class="trio">
          <div><b>${known}</b>외운 단어</div>
          <div><b>${learning}</b>외우는 중</div>
          <div><b>${deck.cards.length - known - learning}</b>아직 안 본 단어</div>
        </div>
      </div>`;
    $('#switchDeck').onclick = () => openOnboard(true);
    $('#checkIn').onclick = () => { if (checkIn(false)) renderHome(true); };
    const s = $('#startStudy'); if (s) s.onclick = () => openStudy();
    const m = $('#moreStudy'); if (m) m.onclick = () => openStudy(10);
    const o = $('#otherDeck'); if (o) o.onclick = () => openOnboard(true);
  }

  // ---------- 단어장 고르기 ----------
  function openOnboard(isSwitch) {
    const el = $('#onboard');
    const list = decks().map((d) => {
      const seen = d.cards.filter((c) => progress[c.id]).length;
      const sel = isSwitch && d.id === settings.deck;
      return `<button class="deckrow${sel ? ' sel' : ''}" data-deck="${d.id}" ${d.cards.length ? '' : 'disabled'}>
        <span class="deckrow-main"><b>${esc(d.name)}</b><span>${esc(d.desc)}${d.id === 'mine' && !d.cards.length ? ' · 단어장 탭에서 넣을 수 있어요' : ''}</span></span>
        <span class="deckrow-cnt">${sel ? '<em>공부 중</em>' : ''}${d.cards.length ? (seen ? `${seen} / ${d.cards.length}` : `${d.cards.length}개`) : '비어 있음'}</span>
      </button>`;
    }).join('');
    el.innerHTML = `<div class="sheet-in">
      <h1 class="headline">어떤 단어를<br><mark>외울까요?</mark></h1>
      <p class="sub">고른 단어장에서 매일 새 단어 ${settings.daily}개가 카드로 나와요. 외운 단어는 잊을 때쯤 다시 꺼내 드려요.</p>
      <div class="rows">${list}</div>
      ${isSwitch ? '<button class="text-btn" id="closeOnboard" style="justify-self:start">그대로 둘게요</button>' : ''}
    </div>`;
    el.hidden = false;
    el.scrollTop = 0;
    $$('.deckrow', el).forEach((b) => b.addEventListener('click', () => {
      settings.deck = b.dataset.deck; save();
      el.hidden = true; go('home');
    }));
    const c = $('#closeOnboard'); if (c) c.onclick = () => { el.hidden = true; };
  }

  // ---------- 학습 ----------
  let S = null;
  function openStudy(extraNew) {
    const q = buildQueue(settings.deck, extraNew);
    if (!q.length) { toast('이 단어장은 끝까지 봤어요. 다른 단어장을 골라 보세요'); return; }
    S = { queue: q, idx: 0, res: [0, 0, 0], flipped: false, dir: 'en' };
    $('#study').hidden = false;
    document.body.style.overflow = 'hidden';
    pushOverlay('study');
    renderCard();
  }
  function closeStudy(fromPop) {
    $('#study').hidden = true;
    document.body.style.overflow = '';
    if (!fromPop) popOverlay();
    if (S && S.idx > 0) checkIn(true);
    S = null;
    try { speechSynthesis.cancel(); } catch (e) { /* noop */ }
    go(currentView);
  }
  const reEsc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  function highlight(ex, word) {
    const re = new RegExp('(' + reEsc(esc(word)).replace(/\s+/g, '\\s+') + '\\w*)', 'i');
    return esc(ex).replace(re, '<mark>$1</mark>');
  }
  const sayBtn = (text, label) => `<button class="say" data-say="${esc(text)}" aria-label="${label}">${ICON.speak}</button>`;
  function renderCard() {
    const el = $('#study');
    if (S.idx >= S.queue.length) return renderDone();
    const c = S.queue[S.idx];
    S.flipped = false;
    S.dir = settings.dir === 'mix' ? (Math.random() < 0.5 ? 'en' : 'ko') : settings.dir;
    const pct = Math.round((S.idx / S.queue.length) * 100);
    const st = status(c);
    const long = c.en.length > 13 ? ' long' : '';
    const front = S.dir === 'en'
      ? `<div class="w${long}">${esc(c.en)}</div>${sayBtn(c.en, '발음 듣기')}<span class="hint">눌러서 뜻 보기</span>`
      : `<div class="w ko">${esc(c.ko)}</div><span class="hint">영어로 떠올린 뒤 눌러서 확인</span>`;
    el.innerHTML = `
      <div class="st-in st-top">
        <button class="x" id="closeStudy" aria-label="학습 닫기">${ICON.close}</button>
        <div class="st-line" aria-hidden="true"><i style="width:${pct}%"></i></div>
        <span class="st-n">${S.idx + 1} / ${S.queue.length}</span>
      </div>
      <div class="stage">
        <div class="fcard" id="flip" role="button" tabindex="0" aria-label="카드 뒤집기">
          <i class="under u2"></i><i class="under u1"></i>
          <div class="flip-in">
            <div class="face front"><span class="tag">${STATUS_LABEL[st]}</span>${front}</div>
            <div class="face back">
              <div class="w sm">${esc(c.en)}</div>
              <div class="mean">${esc(c.ko)}</div>
              ${c.ex ? `<div class="ex">${highlight(c.ex, c.en)}</div>` : ''}
              ${sayBtn(c.ex || c.en, c.ex ? '예문 듣기' : '발음 듣기')}
            </div>
          </div>
        </div>
      </div>
      <div class="st-in st-actions" id="actions"><button class="cta" id="showBtn">뜻 보기</button></div>`;
    $('#closeStudy').onclick = () => closeStudy(false);
    $('#flip').onclick = (e) => {
      const sp = e.target.closest('.say');
      if (sp) { speak(sp.dataset.say); return; }
      flip();
    };
    $('#showBtn').onclick = flip;
    if (settings.autoSpeak && S.dir === 'en') speak(c.en);
  }
  function flip() {
    if (!S || S.flipped) return;
    S.flipped = true;
    const c = S.queue[S.idx];
    $('#flip').classList.add('on');
    if (settings.autoSpeak && S.dir === 'ko') speak(c.en);
    $('#actions').innerHTML = `<div class="oxs">
      <button class="ox ox-x" data-g="0">${ICON.x}<b>모름</b><small>곧 다시</small></button>
      <button class="ox ox-t" data-g="1">${ICON.tri}<b>애매</b><small>${preview(c, 1)}</small></button>
      <button class="ox ox-o" data-g="2">${ICON.o}<b>알아요</b><small>${preview(c, 2)}</small></button>
    </div>`;
    $$('.ox').forEach((b) => b.addEventListener('click', () => doGrade(+b.dataset.g)));
  }
  function doGrade(g) {
    if (!S || !S.flipped) return;
    const c = S.queue[S.idx];
    grade(c, g);
    S.res[g]++;
    if (g === 0) S.queue.splice(Math.min(S.queue.length, S.idx + 4), 0, c);
    S.idx++;
    renderCard();
  }
  function renderDone() {
    const el = $('#study');
    const att = checkIn(true);
    const goalHit = hasDay(stats.goal, today());
    el.innerHTML = `
      <div class="st-in st-top"><button class="x" id="closeStudy" aria-label="학습 닫기">${ICON.close}</button><div class="st-line" aria-hidden="true"><i style="width:100%"></i></div><span class="st-n">끝</span></div>
      <div class="stage"><div class="fin">
        ${BIG_CHECK}
        <h1 class="headline">${goalHit ? `오늘 ${settings.daily}개 <mark>끝!</mark>` : `카드 ${S.idx}장 <mark>끝!</mark>`}</h1>
        <div class="fin-nums"><span>알아요 <b>${S.res[2]}</b></span><span>애매 <b>${S.res[1]}</b></span><span>모름 <b>${S.res[0]}</b></span></div>
        <p class="sub">${streakOf(stats.studied)}일째 이어서 공부 중${att ? ` · 출석 +${att}점` : ''}</p>
      </div></div>
      <div class="st-in st-actions">
        <button class="cta" id="homeBtn">닫기</button>
        <button class="line" id="quizBtn">방금 본 단어로 퀴즈</button>
      </div>`;
    $('#closeStudy').onclick = () => closeStudy(false);
    $('#homeBtn').onclick = () => closeStudy(false);
    $('#quizBtn').onclick = () => { closeStudy(false); go('quiz'); startQuiz(); };
  }
  document.addEventListener('keydown', (e) => {
    if ($('#study').hidden || !S) return;
    if (e.target.closest('input, textarea, .say, .ox, .x')) return;
    if ((e.key === ' ' || e.key === 'Enter') && !S.flipped) { e.preventDefault(); flip(); }
    else if (S.flipped && ['1', '2', '3'].includes(e.key)) doGrade(+e.key - 1);
    else if (e.key === 'Escape') closeStudy(false);
  });

  // ---------- 퀴즈 ----------
  let quizMode = 'choice';
  let Q = null;
  const NUMS = ['1', '2', '3', '4'];
  function renderQuizStart() {
    const deck = deckById(settings.deck);
    const seen = deck.cards.filter((c) => progress[c.id]).length;
    $('#view-quiz').innerHTML = `<div class="page">
      <h1 class="headline">10문제로<br><mark>확인해 볼까요?</mark></h1>
      <div class="tabs" role="group" aria-label="퀴즈 종류">
        <button data-mode="choice" aria-pressed="${quizMode === 'choice'}">뜻 고르기</button>
        <button data-mode="spell" aria-pressed="${quizMode === 'spell'}">철자 쓰기</button>
      </div>
      <p class="note">${esc(deck.name)}에서 ${seen >= 4 ? `지금까지 본 단어 ${seen}개 중에` : '아직 본 단어가 적어서 전체 단어 중에'} 문제를 내요. 맞히면 +${POINTS.quiz}점, 틀린 단어는 오늘 복습에 다시 나와요.</p>
      <button class="cta" id="quizStart" ${deck.cards.length < 4 ? 'disabled' : ''}>퀴즈 시작</button>
    </div>`;
    $$('#view-quiz .tabs button').forEach((b) => b.addEventListener('click', () => { quizMode = b.dataset.mode; renderQuizStart(); }));
    $('#quizStart').onclick = startQuiz;
  }
  function startQuiz() {
    const pool = deckById(settings.deck).cards;
    const seen = pool.filter((c) => progress[c.id]);
    const base = seen.length >= 4 ? seen : pool;
    Q = { items: shuffle(base).slice(0, 10), i: 0, score: 0, wrong: [], answered: false, pool };
    renderQ();
  }
  function renderQ() {
    const view = $('#view-quiz');
    if (Q.i >= Q.items.length) return renderQuizEnd();
    const c = Q.items[Q.i];
    Q.answered = false;
    const head = `<div class="st-top"><div class="st-line" aria-hidden="true"><i style="width:${(Q.i / Q.items.length) * 100}%"></i></div><span class="st-n">${Q.i + 1} / ${Q.items.length}</span></div>`;
    if (quizMode === 'choice') {
      const others = shuffle(Q.pool.filter((x) => x.id !== c.id && x.ko !== c.ko)).slice(0, 3);
      const opts = shuffle([c, ...others]);
      view.innerHTML = `<div class="page">${head}
        <div class="q"><div class="w${c.en.length > 13 ? ' long' : ''}">${esc(c.en)}</div><button class="say on-paper" data-say="${esc(c.en)}" aria-label="발음 듣기">${ICON.speak}</button></div>
        <div class="opts">${opts.map((o, i) => `<button class="opt" data-id="${esc(o.id)}"><span class="no">${NUMS[i]}</span><span>${esc(o.ko)}</span></button>`).join('')}</div>
        <div class="fb" id="fb" aria-live="polite"></div>
        <button class="line" id="nextQ" hidden>다음 문제</button>
      </div>`;
      if (settings.autoSpeak) speak(c.en);
      $$('.opt', view).forEach((b) => b.addEventListener('click', () => answer(b.dataset.id === c.id, b)));
    } else {
      const hint = `${c.en[0]}로 시작 · ${c.en.replace(/\s/g, '').length}글자`;
      view.innerHTML = `<div class="page">${head}
        <div class="q"><div class="mean">${esc(c.ko)}</div>${c.ex ? `<div class="ex">${esc(c.ex.replace(new RegExp(reEsc(c.en) + '\\w*', 'ig'), '_____'))}</div>` : ''}<p class="note">${esc(hint)}</p></div>
        <form id="spellForm" class="stack" autocomplete="off">
          <input class="under-input spell" id="spellIn" autocapitalize="none" autocorrect="off" spellcheck="false" aria-label="영어 단어 입력" placeholder="영어로 쓰기">
          <button class="cta" type="submit" id="spellBtn">확인</button>
        </form>
        <div class="fb" id="fb" aria-live="polite"></div>
        <button class="line" id="nextQ" hidden>다음 문제</button>
      </div>`;
      const inp = $('#spellIn');
      setTimeout(() => inp.focus(), 50);
      $('#spellForm').addEventListener('submit', (e) => {
        e.preventDefault();
        if (Q.answered) return;
        const norm = (s) => s.toLowerCase().replace(/[\s\-.]/g, '');
        answer(norm(inp.value) === norm(c.en), null);
        inp.disabled = true; $('#spellBtn').hidden = true;
      });
    }
    $$('.say', view).forEach((b) => b.addEventListener('click', () => speak(b.dataset.say)));
    $('#nextQ').onclick = () => { Q.i++; renderQ(); };
  }
  function answer(ok, btn) {
    if (Q.answered) return;
    Q.answered = true;
    const c = Q.items[Q.i];
    const fb = $('#fb');
    if (ok) { Q.score++; addPoints(POINTS.quiz); fb.className = 'fb ok'; fb.textContent = `정답 +${POINTS.quiz}점`; }
    else {
      Q.wrong.push(c);
      fb.className = 'fb no';
      if (progress[c.id]) { progress[c.id].d = today(); progress[c.id].b = Math.max(0, progress[c.id].b - 1); }
    }
    if (btn) {
      let rightNo = '';
      $$('.opt').forEach((b) => { b.disabled = true; if (b.dataset.id === c.id) { b.classList.add('right'); rightNo = $('.no', b).textContent; } });
      if (!ok) { btn.classList.add('wrong'); fb.textContent = `정답은 ${rightNo}번이에요`; }
    } else if (!ok) fb.textContent = `정답은 ${c.en}`;
    save();
    if (ok && quizMode === 'choice') setTimeout(() => { if (Q && Q.answered && Q.items[Q.i] === c) { Q.i++; renderQ(); } }, 800);
    else $('#nextQ').hidden = false;
    if (ok && quizMode === 'spell') speak(c.en);
  }
  function renderQuizEnd() {
    const view = $('#view-quiz');
    view.innerHTML = `<div class="page">
      <div class="score">${Q.score}<small> / ${Q.items.length}</small></div>
      <p class="sub">${Q.score === Q.items.length ? '전부 맞혔어요.' : `${Q.wrong.length}개는 오늘 복습에 다시 나와요.`} +${Q.score * POINTS.quiz}점 받았어요.</p>
      ${Q.wrong.length ? `<ul class="wrongs">${Q.wrong.map((w) => `<li><b>${esc(w.en)}</b><span>${esc(w.ko)}</span></li>`).join('')}</ul>` : ''}
      <button class="cta" id="again">한 번 더</button>
      <button class="line" id="toHome">오늘 화면으로</button>
    </div>`;
    $('#again').onclick = startQuiz;
    $('#toHome').onclick = () => go('home');
    checkIn(true);
  }

  // ---------- 단어장 ----------
  let wFilter = null;
  let wQuery = '';
  function renderWords() {
    if (!wFilter) wFilter = settings.deck;
    const ds = decks();
    const view = $('#view-words');
    view.innerHTML = `<div class="page" style="gap:0">
      <input class="under-input search" id="wq" type="search" placeholder="단어나 뜻으로 찾기" value="${esc(wQuery)}" aria-label="단어 검색">
      <div class="tabs" role="group" aria-label="단어장 선택" style="margin-top:8px">
        <button data-f="all" aria-pressed="${wFilter === 'all'}">전체</button>
        ${ds.map((d) => `<button data-f="${d.id}" aria-pressed="${wFilter === d.id}">${esc(d.name)}<small>${d.cards.length}</small></button>`).join('')}
      </div>
      <details class="fold" id="addBox">
        <summary>단어 직접 넣기</summary>
        <form id="addForm" class="fold-body" autocomplete="off">
          <div class="field"><label for="addEn">영어</label><input class="under-input" id="addEn" required autocapitalize="none" spellcheck="false" placeholder="lab dip"></div>
          <div class="field"><label for="addKo">뜻</label><input class="under-input" id="addKo" required placeholder="랩딥 (색상 시험 염색)"></div>
          <div class="field"><label for="addEx">예문 (없어도 돼요)</label><input class="under-input" id="addEx" placeholder="Please send the lab dips by Monday."></div>
          <button class="cta" type="submit">내 단어장에 넣기</button>
        </form>
      </details>
      <details class="fold" id="importBox">
        <summary>엑셀에서 한꺼번에 넣기</summary>
        <div class="fold-body">
          <p class="note">엑셀에서 영어, 뜻, 예문 순서로 된 칸을 복사해서 붙여 넣으세요. 한 줄에 한 단어, 예문은 없어도 돼요.</p>
          <textarea class="box" id="importText" aria-label="붙여 넣을 단어" placeholder="sample	샘플, 견본	We will send the sample.&#10;quotation	견적서"></textarea>
          <div class="two">
            <button class="cta" id="importBtn" style="min-height:52px;font-size:15px">붙여 넣은 단어 넣기</button>
            <label class="line" for="importFile">CSV 파일 고르기</label>
          </div>
          <input type="file" id="importFile" accept=".csv,.txt,.tsv,text/csv,text/plain" hidden>
        </div>
      </details>
      <ul class="wlist" id="wlist"></ul>
    </div>`;
    $('#wq').addEventListener('input', (e) => { wQuery = e.target.value; renderWordList(); });
    $$('.tabs button', view).forEach((b) => b.addEventListener('click', () => { wFilter = b.dataset.f; renderWords(); }));
    const cur = $('.tabs button[aria-pressed="true"]', view);
    if (cur) cur.parentElement.scrollLeft = Math.max(0, cur.offsetLeft - 20);
    $('#addForm').addEventListener('submit', (e) => {
      e.preventDefault();
      const word = $('#addEn').value.trim();
      const n = addCustom([{ en: word, ko: $('#addKo').value, ex: $('#addEx').value }]);
      if (n) { toast(`${word} 넣었어요`); wFilter = 'mine'; renderWords(); $('#addBox').open = true; $('#addEn').focus(); }
      else toast('이미 내 단어장에 있어요');
    });
    $('#importBtn').onclick = () => importText($('#importText').value);
    $('#importFile').addEventListener('change', (e) => {
      const f = e.target.files[0]; if (!f) return;
      const r = new FileReader();
      r.onload = () => importText(String(r.result));
      r.readAsText(f, 'utf-8');
    });
    renderWordList();
  }
  function renderWordList() {
    const ul = $('#wlist');
    const q = wQuery.trim().toLowerCase();
    let list = cardsOf(wFilter);
    if (q) list = list.filter((c) => c.en.toLowerCase().includes(q) || c.ko.includes(q));
    if (!list.length) {
      ul.innerHTML = `<li class="empty">${wFilter === 'mine' && !q ? '아직 넣은 단어가 없어요. 위에서 넣어 보세요.' : '찾는 단어가 없어요.'}</li>`;
      return;
    }
    ul.innerHTML = list.map((c) => {
      const st = status(c);
      return `<li>
        <div class="w-main"><div class="w-en">${esc(c.en)}</div><div class="w-ko">${esc(c.ko)}</div></div>
        ${st === 'new' ? '' : `<span class="st ${st}">${STATUS_LABEL[st]}</span>`}
        <button class="ic" data-say="${esc(c.en)}" aria-label="${esc(c.en)} 발음 듣기">${ICON.speak}</button>
        ${c.deck === 'mine' ? `<button class="ic" data-del="${esc(c.id)}" aria-label="${esc(c.en)} 지우기">${ICON.trash}</button>` : ''}
      </li>`;
    }).join('');
    $$('[data-say]', ul).forEach((b) => b.addEventListener('click', () => speak(b.dataset.say)));
    $$('[data-del]', ul).forEach((b) => b.addEventListener('click', () => {
      if (b.dataset.armed) {
        custom = custom.filter((c) => c.id !== b.dataset.del);
        delete progress[b.dataset.del];
        save(); renderWords(); toast('지웠어요');
      } else {
        b.dataset.armed = '1'; b.style.color = 'var(--red)';
        toast('한 번 더 누르면 지워져요');
        setTimeout(() => { if (b.isConnected) { delete b.dataset.armed; b.style.color = ''; } }, 3000);
      }
    }));
  }
  function addCustom(items) {
    let n = 0;
    items.forEach(({ en, ko, ex }) => {
      en = (en || '').trim(); ko = (ko || '').trim(); ex = (ex || '').trim();
      if (!en || !ko) return;
      const id = 'mine:' + en.toLowerCase();
      if (custom.some((c) => c.id === id)) return;
      custom.push({ id, en, ko, ex });
      n++;
    });
    if (n) save();
    return n;
  }
  function parseCSVLine(line) {
    const out = []; let cur = ''; let q = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (q) { if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch; }
      else if (ch === '"') q = true;
      else if (ch === ',') { out.push(cur); cur = ''; }
      else cur += ch;
    }
    out.push(cur);
    return out.map((s) => s.trim());
  }
  function importText(text) {
    const rows = [];
    String(text || '').replace(/^﻿/, '').split(/\r?\n/).forEach((line) => {
      if (!line.trim()) return;
      let f;
      if (line.includes('\t')) f = line.split('\t').map((s) => s.trim());
      else if (line.includes('|')) f = line.split('|').map((s) => s.trim());
      else {
        f = parseCSVLine(line);
        if (f.length > 3 || (f.length === 3 && /[가-힣]/.test(f[2]))) {
          // 뜻 안에 쉼표가 있는 경우: 마지막 칸이 영어 문장이면 예문으로 봅니다
          const last = f[f.length - 1];
          if (/[가-힣]/.test(last)) f = [f[0], f.slice(1).join(', ')];
          else f = [f[0], f.slice(1, -1).join(', '), last];
        }
      }
      if (/^(word|english|영어|단어)$/i.test(f[0])) return;
      rows.push({ en: f[0], ko: f[1], ex: f[2] });
    });
    const n = addCustom(rows);
    toast(n ? `${n}개를 내 단어장에 넣었어요` : '새로 넣을 단어가 없어요. 형식을 확인해 주세요');
    if (n) { wFilter = 'mine'; renderWords(); }
  }

  // ---------- 설정 ----------
  let installEvt = null;
  window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); installEvt = e; if (currentView === 'settings') renderSettings(); });
  const isStandalone = () => NATIVE || window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
  const isIOS = () => /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  let resetArmed = false;
  function renderSettings() {
    const hours = Array.from({ length: 17 }, (_, i) => i + 6);
    const hourLabel = (h) => `${h < 12 ? '오전' : '오후'} ${h <= 12 ? h : h - 12}시`;
    let installHtml = '';
    if (!isStandalone()) {
      installHtml = `<div class="grp stack">
        <h2>휴대폰에 설치</h2>
        ${installEvt ? '<button class="cta" id="installBtn">홈 화면에 설치하기</button>' : ''}
        ${isIOS()
          ? '<ol class="steps"><li>Safari 아래쪽 공유 버튼을 누르세요</li><li>홈 화면에 추가를 고르세요</li></ol>'
          : '<ol class="steps"><li>Chrome 오른쪽 위 메뉴(⋮)를 누르세요</li><li>앱 설치 또는 홈 화면에 추가를 고르세요</li></ol>'}
        <p class="note">설치하면 인터넷 없이도 열려요.</p>
      </div><div class="band"></div>`;
    }
    $('#view-settings').innerHTML = `
      ${installHtml}
      <div class="grp">
        <h2>공부</h2>
        <div class="set"><div><b>단어장</b></div><button class="sel" id="setDeck">${esc(deckById(settings.deck).name)}</button></div>
        <div class="set"><div><b>하루 새 단어</b></div>
          <select class="sel" id="setDaily" aria-label="하루 새 단어 수">${[10, 15, 20, 30, 40].map((n) => `<option value="${n}" ${settings.daily === n ? 'selected' : ''}>${n}개</option>`).join('')}</select></div>
        <div class="set"><div><b>카드 앞면</b></div>
          <select class="sel" id="setDir" aria-label="카드 앞면"><option value="en" ${settings.dir === 'en' ? 'selected' : ''}>영어 먼저</option><option value="ko" ${settings.dir === 'ko' ? 'selected' : ''}>뜻 먼저</option><option value="mix" ${settings.dir === 'mix' ? 'selected' : ''}>섞어서</option></select></div>
        <div class="set"><div><b>발음 자동으로 듣기</b></div><label class="switch"><input type="checkbox" id="setSpeak" ${settings.autoSpeak ? 'checked' : ''} aria-label="발음 자동으로 듣기"><span></span></label></div>
        <div class="set"><div><b>읽는 속도</b></div>
          <select class="sel" id="setRate" aria-label="읽는 속도">${[[0.7, '천천히'], [0.9, '보통'], [1.1, '빠르게']].map(([v, l]) => `<option value="${v}" ${settings.rate === v ? 'selected' : ''}>${l}</option>`).join('')}</select></div>
      </div>
      <div class="band"></div>
      <div class="grp">
        <h2>알림</h2>
        <div class="set"><div><b>공부 안 한 날 알려 주기</b><p>그날 공부했으면 울리지 않아요</p></div><label class="switch"><input type="checkbox" id="setRemind" ${settings.remind ? 'checked' : ''} aria-label="공부 안 한 날 알려 주기"><span></span></label></div>
        <div class="set"><div><b>알림 시간</b></div>
          <select class="sel" id="setHour" aria-label="알림 시간" ${settings.remind ? '' : 'disabled'}>${hours.map((h) => `<option value="${h}" ${settings.remindHour === h ? 'selected' : ''}>${hourLabel(h)}</option>`).join('')}</select></div>
        <div class="set"><div><b>알림 테스트</b><p>${NATIVE ? '3초 뒤에 알림이 와요' : '브라우저에서는 늦거나 안 올 수 있어요'}</p></div><button class="sel" id="testNotify">보내 보기</button></div>
      </div>
      <div class="band"></div>
      <div class="grp stack">
        <h2>기록 옮기기</h2>
        <p class="note">공부 기록과 점수는 이 휴대폰에만 있어요. 휴대폰을 바꾸기 전에 백업 코드를 복사해 두세요.</p>
        <div class="two"><button class="line" id="copyBackup">백업 코드 복사</button><button class="line" id="showRestore">백업 불러오기</button></div>
        <div id="restoreBox" class="stack" hidden>
          <textarea class="box" id="restoreText" aria-label="백업 코드" placeholder="복사해 둔 백업 코드를 붙여 넣으세요"></textarea>
          <button class="cta" id="restoreBtn">이 백업으로 바꾸기</button>
        </div>
        <button class="text-btn" id="resetBtn" style="justify-self:start">공부 기록 전부 지우기</button>
      </div>`;
    const ib = $('#installBtn'); if (ib) ib.onclick = async () => { installEvt.prompt(); await installEvt.userChoice.catch(() => {}); installEvt = null; renderSettings(); };
    $('#setDeck').onclick = () => openOnboard(true);
    $('#setDaily').onchange = (e) => { settings.daily = +e.target.value; save(); toast(`하루 ${settings.daily}개로 바꿨어요`); };
    $('#setDir').onchange = (e) => { settings.dir = e.target.value; save(); };
    $('#setSpeak').onchange = (e) => { settings.autoSpeak = e.target.checked; save(); };
    $('#setRate').onchange = (e) => { settings.rate = +e.target.value; save(); speak('Hello, nice to meet you.'); };
    $('#setRemind').onchange = async (e) => {
      settings.remind = e.target.checked; save();
      $('#setHour').disabled = !settings.remind;
      if (settings.remind) await requestNotify();
      syncReminder();
    };
    $('#setHour').onchange = (e) => { settings.remindHour = +e.target.value; save(); toast(`${hourLabel(settings.remindHour)}에 알려 드릴게요`); };
    $('#testNotify').onclick = async () => {
      const ok = await requestNotify();
      if (!ok) return;
      const LN = plugin('LocalNotifications');
      if (LN) await LN.schedule({ notifications: [{ id: 999, title: '단어카드', body: '알림이 잘 와요. 공부 안 한 날 이렇게 알려 드릴게요.', schedule: { at: new Date(Date.now() + 3000) } }] }).catch(() => {});
      else showWebNotice();
      syncReminder();
      toast('곧 알림이 와요');
    };
    $('#copyBackup').onclick = async () => {
      const code = JSON.stringify({ app: 'wordcard', v: 1, settings, progress, custom, stats });
      try { await navigator.clipboard.writeText(code); toast('복사했어요. 메모장이나 카톡 나에게 보내기에 붙여 두세요'); }
      catch (e) { $('#restoreBox').hidden = false; $('#restoreText').value = code; $('#restoreText').select(); toast('아래 코드를 길게 눌러 복사하세요'); }
    };
    $('#showRestore').onclick = () => { $('#restoreBox').hidden = !$('#restoreBox').hidden; $('#restoreText').value = ''; };
    $('#restoreBtn').onclick = () => {
      try {
        const d = JSON.parse($('#restoreText').value);
        if (d.app !== 'wordcard') throw new Error('bad');
        settings = Object.assign({}, DEFAULT_SETTINGS, d.settings); progress = d.progress || {}; custom = d.custom || [];
        stats = Object.assign({ pts: 0, attend: [], studied: [], goal: [], newBy: {} }, d.stats);
        save(); renderPoints(); toast('백업을 불러왔어요'); go('home');
      } catch (e) { toast('백업 코드가 맞지 않아요'); }
    };
    $('#resetBtn').onclick = (e) => {
      if (!resetArmed) { resetArmed = true; e.target.textContent = '정말 지울까요? 한 번 더 누르세요'; e.target.style.color = 'var(--red)'; setTimeout(() => { resetArmed = false; if (e.target.isConnected) { e.target.textContent = '공부 기록 전부 지우기'; e.target.style.color = ''; } }, 4000); return; }
      resetArmed = false;
      progress = {}; stats = { pts: 0, attend: [], studied: [], goal: [], newBy: {} };
      save(); renderPoints(); toast('기록을 지웠어요'); go('home');
    };
  }

  // ---------- 시작 ----------
  if (!NATIVE && 'serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost')) {
    try { navigator.serviceWorker.register('sw.js').catch(() => {}); } catch (e) { /* 미리보기 환경 */ }
  }
  // 앱으로 돌아오면(다음 날 포함) 화면과 알림 예약을 새로 고칩니다
  document.addEventListener('visibilitychange', () => {
    if (document.hidden) return;
    if (currentView === 'home' && $('#study').hidden) renderHome();
    syncReminder();
  });

  renderPoints();
  if (!settings.deck || !decks().some((d) => d.id === settings.deck)) { settings.deck = 'basic'; openOnboard(false); }
  go('home');
  syncReminder();
})();
