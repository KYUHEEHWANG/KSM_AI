using System;
using System.Runtime.InteropServices;

namespace KSM_SolidEdge
{
    public static class SleepModeController
    {
        // Windows API 선언
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern uint SetThreadExecutionState(uint esFlags);

        // 플래그 설정
        private const uint ES_CONTINUOUS = 0x80000000;
        private const uint ES_SYSTEM_REQUIRED = 0x00000001;
        private const uint ES_DISPLAY_REQUIRED = 0x00000002;

        /// <summary>
        /// PC가 절전 모드로 들어가는 것을 방지합니다. (화면 유지 포함)
        /// </summary>
        public static void PreventSleep()
        {
            // 시스템 요구 및 디스플레이 요구 상태를 지속적으로 유지
            SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED);
        }

        /// <summary>
        /// 설정된 절전 방지 모드를 해제하고 시스템 설정에 따르도록 복구합니다.
        /// </summary>
        public static void AllowSleep()
        {
            SetThreadExecutionState(ES_CONTINUOUS);
        }
    }
}