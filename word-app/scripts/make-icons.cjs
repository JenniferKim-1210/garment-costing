// 앱 아이콘과 시작 화면 이미지를 다시 그립니다. (Playwright 필요: node scripts/make-icons.cjs)
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
// 512 기준 좌표: 겹친 링 단어카드
const art = `<g transform="rotate(9 256 270)"><rect x="150" y="128" width="212" height="268" rx="22" fill="#FFA8C8"/></g>
<g transform="rotate(-7 256 270)"><rect x="150" y="128" width="212" height="268" rx="22" fill="#FFD43B"/>
<circle cx="256" cy="160" r="13" fill="#111"/>
<text x="256" y="312" text-anchor="middle" font-family="Inter, DejaVu Sans, sans-serif" font-weight="900" font-size="112" letter-spacing="-5" fill="#111">Aa</text></g>
<circle cx="250" cy="138" r="27" fill="none" stroke="#B9B9B5" stroke-width="9"/>`;
const svg = (w, bg, scale) => `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${w}" viewBox="0 0 ${w} ${w}">${bg ? `<rect width="${w}" height="${w}" fill="${bg}"/>` : ''}<g transform="translate(${w / 2} ${w / 2}) scale(${scale}) translate(-256 -262)">${art}</g></svg>`;

fs.writeFileSync(path.join(root, 'icons/icon.svg'), svg(512, '#111111', 1.15).replace(' width="512" height="512"', ''));
const jobs = [
  ['icons/icon-192.png', 192, '#111111', (192 / 512) * 1.15],
  ['icons/icon-512.png', 512, '#111111', 1.15],
  ['icons/apple-touch-icon.png', 180, '#111111', (180 / 512) * 1.15],
  ['assets/icon-only.png', 1024, '#111111', 2.3],
  ['assets/icon-foreground.png', 1024, null, 1.6],
  ['assets/icon-background.png', 1024, '#111111', 1],
  ['assets/splash.png', 2732, '#FFFFFF', 2.2],
  ['assets/splash-dark.png', 2732, '#111111', 2.2]
];
(async () => {
  const b = await chromium.launch();
  for (const [name, w, bg, sc] of jobs) {
    const p = await b.newPage({ viewport: { width: w, height: w } });
    await p.setContent('<style>html,body{margin:0;background:transparent}svg{display:block}</style>' + svg(w, bg, sc));
    await p.screenshot({ path: path.join(root, name), omitBackground: !bg });
    await p.close();
  }
  await b.close();
  console.log('icons done');
})();
