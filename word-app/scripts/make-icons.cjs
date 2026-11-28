// 앱 아이콘과 시작 화면 이미지를 다시 그립니다. (Playwright 필요: node scripts/make-icons.cjs)
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
// 512 기준 좌표: 겹쳐 놓은 카드 세 장
const art = `<rect x="178" y="118" width="156" height="60" rx="30" fill="#3A3A3F"/>
<rect x="160" y="140" width="192" height="80" rx="34" fill="#77777D"/>
<rect x="140" y="166" width="232" height="240" rx="44" fill="#FFFFFF"/>
<text x="256" y="322" text-anchor="middle" font-family="Inter, DejaVu Sans, sans-serif" font-weight="800" font-size="116" letter-spacing="-6" fill="#111113">Aa</text>
<circle cx="336" cy="202" r="13" fill="#2F6BFF"/>`;
const svg = (w, bg, scale) => `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${w}" viewBox="0 0 ${w} ${w}">${bg ? `<rect width="${w}" height="${w}" fill="${bg}"/>` : ''}<g transform="translate(${w / 2} ${w / 2}) scale(${scale}) translate(-256 -262)">${art}</g></svg>`;

fs.writeFileSync(path.join(root, 'icons/icon.svg'), svg(512, '#111113', 1.15).replace(' width="512" height="512"', ''));
const jobs = [
  ['icons/icon-192.png', 192, '#111113', (192 / 512) * 1.15],
  ['icons/icon-512.png', 512, '#111113', 1.15],
  ['icons/apple-touch-icon.png', 180, '#111113', (180 / 512) * 1.15],
  ['assets/icon-only.png', 1024, '#111113', 2.3],
  ['assets/icon-foreground.png', 1024, null, 1.6],
  ['assets/icon-background.png', 1024, '#111113', 1],
  ['assets/splash.png', 2732, '#FFFFFF', 2.0],
  ['assets/splash-dark.png', 2732, '#0E0E10', 2.0]
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
