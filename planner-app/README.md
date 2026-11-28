# 하루플래너 (Haru Planner)

할 일 · 캘린더 · 미리 알림 · 집중 타이머를 한곳에 모은 플래너 앱입니다.
React + TypeScript + Vite로 만든 웹앱(PWA)이며, Capacitor로 감싸 안드로이드 앱(Play Store)으로 낼 수 있습니다.

## 기능

- **할 일**: 오늘 / 내일 / 예정 / 언젠가 구분, 목록(개인·업무·장보기 등), 하위 작업, 메모, 중요도
- **손쉬운 조작**: 오른쪽으로 밀어 완료, 왼쪽으로 밀어 삭제(실행 취소 가능), `⋮⋮`를 끌어 순서·날짜 변경
- **빠른 입력**: "내일 오후 3시 회의"처럼 쓰면 날짜와 시간을 자동으로 읽음, 음성 입력
- **캘린더**: 월 / 주 / 3일 / 일 / 목록 보기, 일정 색상, 겹치는 일정 나란히 표시
- **미리 알림**: 일회성·반복(매일, 평일, 매주, 매월, 매년) 알림
- **집중 모드**: 집중/휴식 타이머, 오늘 완료한 집중 횟수
- **공유**: 목록이나 항목을 텍스트로 공유(장보기 목록을 메신저로 보내기 등)
- **그 밖에**: 한국어/영어, 라이트/다크 테마, 오프라인 동작, JSON 백업·복원

모든 데이터는 기기 안(localStorage)에만 저장됩니다. 서버가 없어서 기기 간 동기화, 실시간 공동 편집,
구글 캘린더 같은 외부 서비스 연동은 아직 없습니다.

## 실행

```bash
npm install
npm run dev      # 개발 서버
npm test         # 날짜·반복·빠른 입력 로직 테스트
npm run build    # dist/ 에 배포용 파일 생성
```

## Play Store에 올리기

안드로이드 빌드는 Android Studio(JDK 21, Android SDK)가 설치된 PC에서 진행합니다.

1. `capacitor.config.ts`의 `appId`를 본인 소유의 고유 ID(예: `kr.내도메인.haruplanner`)로 바꿉니다. **출시 후에는 바꿀 수 없습니다.**
2. 안드로이드 프로젝트 생성 및 동기화
   ```bash
   npm run android:add     # 처음 한 번
   npm run android:sync    # 웹 코드를 바꿀 때마다
   npm run android:open    # Android Studio 열기
   ```
3. `android/app/src/main/AndroidManifest.xml`에 권한 추가
   ```xml
   <uses-permission android:name="android.permission.POST_NOTIFICATIONS" />
   <uses-permission android:name="android.permission.SCHEDULE_EXACT_ALARM" />
   <uses-permission android:name="android.permission.RECEIVE_BOOT_COMPLETED" />
   ```
   (정확한 시각 알림이 필요 없으면 `SCHEDULE_EXACT_ALARM`은 빼도 됩니다. 넣을 경우 Play Console에서 사용 사유를 물어봅니다.)
4. 알림 아이콘 `ic_stat_icon`(흰색 단색 PNG)을 `android/app/src/main/res/drawable/`에 넣고,
   Android Studio의 **Image Asset** 도구로 `public/icon-512.png`를 앱 아이콘으로 등록합니다.
5. **Build > Generate Signed App Bundle**로 `.aab` 파일을 만듭니다. 서명 키(keystore)는 잃어버리면 업데이트를 못 올리니 꼭 백업하세요.
6. [Play Console](https://play.google.com/console) 개발자 계정(등록비 25달러, 1회)에서 앱을 만들고 `.aab`를 올립니다.
   - 개인정보처리방침 URL이 필요합니다. `public/privacy.html`의 빈칸(시행일, 문의 이메일)을 채워 웹에 올린 뒤 그 주소를 넣으세요.
   - 데이터 보안 양식: 이 앱은 데이터를 수집·전송하지 않습니다.
   - 2023년 11월 이후 만든 개인 개발자 계정은 정식 출시 전에 12명 이상이 14일간 참여하는 비공개 테스트가 필요합니다.

## 알아 둘 점

- 웹(PWA)에서는 앱이 열려 있을 때만 알림이 울립니다. 안드로이드 앱에서는 꺼져 있어도 울립니다.
- 음성 입력은 브라우저의 음성 인식 기능을 씁니다. 안드로이드 앱(WebView)에서는 동작하지 않아 안내 문구가 뜹니다.
  앱에서도 쓰려면 `@capacitor-community/speech-recognition` 플러그인을 붙여야 합니다.
- 안드로이드 앱의 알림 예약은 실제 기기에서 확인이 필요합니다(이 저장소에서는 웹 동작만 검증했습니다).
