// cap add android 로 만든 프로젝트에 알림 아이콘과 버전 번호를 넣습니다.
import { cpSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));

cpSync(join(root, 'android-res'), join(root, 'android/app/src/main/res'), { recursive: true });

// 플레이스토어는 올릴 때마다 versionCode 가 커져야 합니다. CI 실행 번호를 씁니다.
const code = parseInt(process.env.VERSION_CODE || '1', 10);
const gradlePath = join(root, 'android/app/build.gradle');
let gradle = readFileSync(gradlePath, 'utf8');
gradle = gradle.replace(/versionCode\s+\d+/, `versionCode ${code}`).replace(/versionName\s+"[^"]*"/, `versionName "${pkg.version}"`);
writeFileSync(gradlePath, gradle);
console.log(`versionCode ${code}, versionName ${pkg.version}`);
