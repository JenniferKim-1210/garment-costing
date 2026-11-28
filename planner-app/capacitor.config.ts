import type { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  // Play Store 등록 전에 본인 소유 도메인 기준의 고유 ID로 바꾸세요. 출시 후에는 변경할 수 없습니다.
  appId: 'com.example.haruplanner',
  appName: '하루플래너',
  webDir: 'dist',
  plugins: {
    LocalNotifications: { smallIcon: 'ic_stat_icon', iconColor: '#0F766E' },
  },
};

export default config;
