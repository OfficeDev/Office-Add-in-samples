using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace DPIHelper
{
    static class DPIHelper
    {
        // Hack until sdk version of windef.h is updated
        public enum DPI_HOSTING_BEHAVIOR
        {
            DPI_HOSTING_BEHAVIOR_INVALID = -1,
            DPI_HOSTING_BEHAVIOR_DEFAULT = 0,
            DPI_HOSTING_BEHAVIOR_MIXED = 1
        }

        [DllImport("SHCore.dll", SetLastError = true)]
        private static extern bool SetProcessDpiAwareness(PROCESS_DPI_AWARENESS awareness);

        [DllImport("SHCore.dll", SetLastError = true)]
        private static extern void GetProcessDpiAwareness(IntPtr hprocess, out PROCESS_DPI_AWARENESS awareness);

        [DllImport("User32.dll", SetLastError = true)]
        private static extern DPI_AWARENESS_CONTEXT SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT awareness);

        [DllImport("User32.dll", SetLastError = true)]
        private static extern DPI_AWARENESS_CONTEXT GetThreadDpiAwarenessContext();

        [DllImport("User32.dll", SetLastError = true)]
        private static extern DPI_AWARENESS_CONTEXT GetWindowDpiAwarenessContext(IntPtr hWnd);

        [DllImport("User32.dll", SetLastError = true)]
        private static extern DPI_AWARENESS GetAwarenessFromDpiAwarenessContext(DPI_AWARENESS_CONTEXT value);

        // DPI_HOSTING_BEHAVIOR WINAPI SetThreadDpiHostingBehavior(_In_ DPI_HOSTING_BEHAVIOR dpiHostingBehavior);
        [DllImport("User32.dll", SetLastError = true)]
        private static extern DPI_HOSTING_BEHAVIOR SetThreadDpiHostingBehavior(DPI_HOSTING_BEHAVIOR dpiHostingBehavior);

        // DPI_HOSTING_BEHAVIOR WINAPI GetThreadDpiHostingBehavior(_In_ HWND hwnd);
        [DllImport("User32.dll", SetLastError = true)]
        private static extern DPI_HOSTING_BEHAVIOR GetThreadDpiHostingBehavior(IntPtr hWnd);

        internal enum PROCESS_DPI_AWARENESS
        {
            Process_DPI_Unaware = 0,
            Process_System_DPI_Aware = 1,
            Process_Per_Monitor_DPI_Aware = 2
        }

        internal enum DPI_AWARENESS_CONTEXT
        {
            DPI_AWARENESS_CONTEXT_UNAWARE = 16,
            DPI_AWARENESS_CONTEXT_SYSTEM_AWARE = 17,
            DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE = 18,
            DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = 34
        }

        internal enum DPI_AWARENESS
        {
            DPI_AWARENESS_INVALID = -1,
            DPI_AWARENESS_UNAWARE = 0,
            DPI_AWARENESS_SYSTEM_AWARE = 1,
            DPI_AWARENESS_PER_MONITOR_AWARE = 2
        }

        public static bool SetDpiAwareness(PROCESS_DPI_AWARENESS awareness)
        {
            return SetProcessDpiAwareness(awareness);
        }

        public static DPI_AWARENESS_CONTEXT SetThreadDpiAwareness(DPI_AWARENESS_CONTEXT awareness)
        {
            return SetThreadDpiAwarenessContext(awareness);
        }

        public static PROCESS_DPI_AWARENESS GetProcessDpi()
        {
            return GetProcessDpi(Process.GetCurrentProcess().Handle);
        }

        public static PROCESS_DPI_AWARENESS GetProcessDpi(IntPtr hprocess)
        {
            PROCESS_DPI_AWARENESS awareness;
            GetProcessDpiAwareness(hprocess, out awareness);
            return awareness;
        }

        public static DPI_AWARENESS GetThreadDpiAwareness()
        {
            DPI_AWARENESS_CONTEXT context = GetThreadDpiAwarenessContext();
            return GetAwarenessFromDpiAwarenessContext(context);
        }

        public static DPI_AWARENESS GetWindowDpiAwareness(IntPtr hWnd)
        {
            DPI_AWARENESS_CONTEXT context = GetWindowDpiAwarenessContext(hWnd);
            return GetAwarenessFromDpiAwarenessContext(context);
        }

        public static DPI_HOSTING_BEHAVIOR SetChildWindowMixedMode(DPI_HOSTING_BEHAVIOR value)
        {
            return SetThreadDpiHostingBehavior(value);
        }

        public static DPI_HOSTING_BEHAVIOR GetChildWindowMixedMode(IntPtr hWnd)
        {
            return GetThreadDpiHostingBehavior(hWnd);
        }

        public static void DebugPrintDPIAwareness(IntPtr hprocess, string message)
        {
            Debug.WriteLine(DPIAwarenessText(hprocess, message));
        }

        public static string DPIAwarenessText(IntPtr hprocess, string message)
        {
            return String.Format("***{0}: Process {1}, Thread {2}", message, GetProcessDpi(hprocess), GetThreadDpiAwareness());
        }

        public static string DPIAwarenessText(string message)
        {
            return DPIAwarenessText(Process.GetCurrentProcess().Handle, message);
        }

    }
}
