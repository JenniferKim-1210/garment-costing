// 앱 파일을 www/ 폴더로 복사합니다. (안드로이드 앱은 www/ 를 그대로 담습니다)
import { cpSync, mkdirSync, rmSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const out = join(root, 'www');
rmSync(out, { recursive: true, force: true });
mkdirSync(out, { recursive: true });
for (const f of ['index.html', 'app.js', 'words.js', 'sw.js', 'manifest.webmanifest', 'privacy.html', 'icons']) {
  cpSync(join(root, f), join(out, f), { recursive: true });
}
console.log('www/ 준비 완료');
