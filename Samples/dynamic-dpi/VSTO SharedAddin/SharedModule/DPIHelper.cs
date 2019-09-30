// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;

namespace SharedModule
{
    public static class DPIHelper
    {
        public enum WinMessages : int
        {
            WM_DPICHANGED = 0x02E0,
            WM_GETMINMAXINFO = 0x0024,
            WM_SIZE = 0x0005,
            WM_WINDOWPOSCHANGING = 0x0046,
            WM_WINDOWPOSCHANGED = 0x0047,
        }
        
        // Was getting AVs using delegate function, switching to declared extern, which should be available back to Win 8.1
        [DllImport("SHCore.dll", SetLastError = true, EntryPoint = "GetProcessDpiAwareness")]
        private static extern void _GetProcessDpiAwareness(IntPtr hprocess, out DPI_AWARENESS awareness);

        [DllImport("SHCore.dll", SetLastError = true, EntryPoint = "SetProcessDpiAwareness")]
        private static extern IntPtr _SetProcessDpiAwareness(DPI_AWARENESS awareness);


        #region Dynamic Dpi delegates
        // User32.dll
        private delegate DPI_AWARENESS_CONTEXT _SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT awareness);
        private delegate DPI_AWARENESS_CONTEXT _GetThreadDpiAwarenessContext();
        private delegate DPI_AWARENESS_CONTEXT _GetWindowDpiAwarenessContext(IntPtr hWnd);
        private delegate DPI_AWARENESS _GetAwarenessFromDpiAwarenessContext(DPI_AWARENESS_CONTEXT value);
        private delegate bool _AreDpiAwarenessContextsEqual(DPI_AWARENESS_CONTEXT valueA, DPI_AWARENESS_CONTEXT valueB);
        private delegate DPI_HOSTING_BEHAVIOR _SetThreadDpiHostingBehavior(DPI_HOSTING_BEHAVIOR dpiHostingBehavior);
        private delegate DPI_HOSTING_BEHAVIOR _GetThreadDpiHostingBehavior();
        private delegate DPI_HOSTING_BEHAVIOR _GetWindowDpiHostingBehavior(IntPtr hWnd);
        #endregion

        #region Other External Apis
        [DllImport("user32.dll", EntryPoint = "GetDpiForWindow")]
        private static extern uint _GetDpiForWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, ref RECT rect);

        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder className, int charCount);

        [DllImport("user32.dll", EntryPoint = "GetAncestor")]
        private static extern IntPtr _GetAncestor(IntPtr hWnd, uint gaFlags);
        [DllImport("kernel32.dll")]
        public static extern IntPtr LoadLibrary(string dllToLoad);

        [DllImport("kernel32.dll")]
        public static extern IntPtr GetProcAddress(IntPtr hModule, string procedureName);

        [DllImport("kernel32.dll")]
        public static extern bool FreeLibrary(IntPtr hModule);
        #endregion

        public struct DPI_AWARENESS_CONTEXT
        {
            private IntPtr value;

            private DPI_AWARENESS_CONTEXT(IntPtr value)
            {
                this.value = value;
            }
            public static implicit operator DPI_AWARENESS_CONTEXT(IntPtr value)
            {
                return new DPI_AWARENESS_CONTEXT(value);
            }

            public static implicit operator IntPtr(DPI_AWARENESS_CONTEXT context)
            {
                return context.value;
            }

            //public static DPI_AWARENESS_CONTEXT operator -(DPI_AWARENESS_CONTEXT context, long value)
            //{
            //    return (IntPtr)(context.value.ToInt64() - value);
            //}
            //public static DPI_AWARENESS_CONTEXT operator -(DPI_AWARENESS_CONTEXT context, int value)
            //{
            //    return (IntPtr)(context.value.ToInt32() - value);
            //}

            //public static bool operator ==(DPI_AWARENESS_CONTEXT context1, DPI_AWARENESS_CONTEXT context2)
            //{
            //    return context1.value == context2;
            //}
            //public static bool operator !=(DPI_AWARENESS_CONTEXT context1, DPI_AWARENESS_CONTEXT context2)
            //{
            //    return context1.value != context2;
            //}

            public static bool operator ==(IntPtr context1, DPI_AWARENESS_CONTEXT context2)
            {
                return AreDpiAwarenessContextsEqual(context1, context2);
            }
            public static bool operator !=(IntPtr context1, DPI_AWARENESS_CONTEXT context2)
            {
                return !AreDpiAwarenessContextsEqual(context1, context2);
            }

            public override bool Equals(object obj)
            {
                return base.Equals(obj);
            }

            public override int GetHashCode()
            {
                return base.GetHashCode();
            }

            public override string ToString()
            {
                if (this.value == DPI_AWARENESS_CONTEXT_UNAWARE)
                {
                    return "Unaware";
                }
                if (this.value == DPI_AWARENESS_CONTEXT_SYSTEM_AWARE)
                {
                    return "System Aware";
                }
                if (this.value == DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE)
                {
                    return "Per Monitor Aware";
                }
                if (this.value == DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)
                {
                    return "Per Monitor Aware V2";
                }
                return "Unknown";
            }
        }

        private static DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_HANDLE = IntPtr.Zero;

        public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_INVALID = IntPtr.Zero;
        public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_UNAWARE = new IntPtr(-1);
        public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_SYSTEM_AWARE = new IntPtr(-2);
		public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE = new IntPtr(-3);
		public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);

		public static DPI_AWARENESS_CONTEXT[] DpiAwarenessContexts =
        {
            DPI_AWARENESS_CONTEXT_UNAWARE,
            DPI_AWARENESS_CONTEXT_SYSTEM_AWARE,
            DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE,
            DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
        };

        public enum DPI_AWARENESS
        {
            DPI_AWARENESS_INVALID = -1,
            DPI_AWARENESS_UNAWARE = 0,
            DPI_AWARENESS_SYSTEM_AWARE = 1,
            DPI_AWARENESS_PER_MONITOR_AWARE = 2
        }

		public static DPI_HOSTING_BEHAVIOR[] DpiHostingBehaviors =
		{
			DPI_HOSTING_BEHAVIOR.DPI_HOSTING_BEHAVIOR_DEFAULT,
			DPI_HOSTING_BEHAVIOR.DPI_HOSTING_BEHAVIOR_MIXED
		};

		public enum DPI_HOSTING_BEHAVIOR
        {
            DPI_HOSTING_BEHAVIOR_INVALID = -1,
            DPI_HOSTING_BEHAVIOR_DEFAULT = 0,
            DPI_HOSTING_BEHAVIOR_MIXED = 1
        }

        public const uint GA_ROOT = 2;

        private static T GetDelegatedFunction<T>(string library, string entrypoint)
        {
            IntPtr pLib = LoadLibrary(library);
            if (pLib == IntPtr.Zero) return default(T);

            IntPtr pProc = GetProcAddress(pLib, entrypoint);
            if (pProc == IntPtr.Zero) return default(T);

            T pWrapped = Marshal.GetDelegateForFunctionPointer<T>(pProc);
            FreeLibrary(pLib);
            return pWrapped;
        }

        public static DPI_AWARENESS_CONTEXT SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT context)
        {
            var pFunction = GetDelegatedFunction<_SetThreadDpiAwarenessContext>("user32.dll", "SetThreadDpiAwarenessContext");
            if (pFunction == null)
                return DPI_AWARENESS_CONTEXT_INVALID;
            return pFunction(context);
        }

        public static DPI_AWARENESS GetProcessDpiAwareness()
        {
            DPI_AWARENESS result = DPI_AWARENESS.DPI_AWARENESS_INVALID;
            IntPtr pid = Process.GetCurrentProcess().Handle;
            try
            {
                _GetProcessDpiAwareness(pid, out result);
            }
            catch (Exception)
            {
            }
            return result;
        }

        public static IntPtr SetProcessDpiAwareness(DPI_AWARENESS awareness)
        {
            return _SetProcessDpiAwareness(awareness);
        }

        public static DPI_AWARENESS_CONTEXT GetThreadDpiAwarenessContext()
        {
            var pFunction = GetDelegatedFunction<_GetThreadDpiAwarenessContext>("user32.dll", "GetThreadDpiAwarenessContext");
            if (pFunction == null)
                return DPI_AWARENESS_CONTEXT_INVALID;
            return pFunction();
        }

        public static DPI_AWARENESS_CONTEXT GetWindowDpiAwarenessContext(IntPtr hWnd)
        {
            var pFunction = GetDelegatedFunction<_GetWindowDpiAwarenessContext>("user32.dll", "GetWindowDpiAwarenessContext");
            if (pFunction == null)
                return DPI_AWARENESS_CONTEXT_INVALID;
            return pFunction(hWnd);
        }

        public static bool AreDpiAwarenessContextsEqual(DPI_AWARENESS_CONTEXT valueA, DPI_AWARENESS_CONTEXT valueB)
        {
            var pFunction = GetDelegatedFunction<_AreDpiAwarenessContextsEqual>("user32.dll", "AreDpiAwarenessContextsEqual");
            if (pFunction == null)
                return false;
            return pFunction(valueA, valueB);
        }

        public static DPI_HOSTING_BEHAVIOR SetThreadDpiHostingBehavior(DPI_HOSTING_BEHAVIOR value)
        {
            var pFunction = GetDelegatedFunction<_SetThreadDpiHostingBehavior>("user32.dll", "SetThreadDpiHostingBehavior");
            if (pFunction == null)
                return DPI_HOSTING_BEHAVIOR.DPI_HOSTING_BEHAVIOR_INVALID;
            return pFunction(value);
        }

        public static DPI_HOSTING_BEHAVIOR GetThreadDpiHostingBehavior()
        {
            var pFunction = GetDelegatedFunction<_GetThreadDpiHostingBehavior>("user32.dll", "GetThreadDpiHostingBehavior");
            if (pFunction == null)
                return DPI_HOSTING_BEHAVIOR.DPI_HOSTING_BEHAVIOR_INVALID;
            return pFunction();
        }

        public static DPI_HOSTING_BEHAVIOR GetWindowDpiHostingBehavior(IntPtr hWnd)
        {
            var pFunction = GetDelegatedFunction<_GetWindowDpiHostingBehavior>("user32.dll", "GetWindowDpiHostingBehavior");
            if (pFunction == null)
                return DPI_HOSTING_BEHAVIOR.DPI_HOSTING_BEHAVIOR_INVALID;
            return pFunction(hWnd);
        }

        public static uint GetDpiForWindow(IntPtr hWnd)
        {
            return _GetDpiForWindow(hWnd);
        }
        public static SizeF GetDpiForWindowSizeF(IntPtr hWnd)
        {
            SizeF scaleFactor = SizeF.Empty;
            scaleFactor.Height =  scaleFactor.Width = (float)_GetDpiForWindow(hWnd);

            return scaleFactor;
        }

        public static RECT GetWindowRectangle(IntPtr hWnd)
        {
            RECT rect = new RECT();
            GetWindowRect(hWnd, ref rect);

            return rect;
        }

        public static IntPtr GetParentWindow(IntPtr hWnd)
        {
            return GetParent(hWnd);
        }

        public static string GetWindowClassName(IntPtr hWnd)
        {
            StringBuilder buff = new StringBuilder(256);
            int retCount = 0;

            retCount = GetClassName(hWnd, buff, 256);

            return buff.ToString();
        }

        public static IntPtr GetAncestor(IntPtr hWnd, uint gaFlags)
        {
            return _GetAncestor(hWnd, gaFlags);
        }

        public static IntPtr FindParentWithClassName(IntPtr hWndChild, string className)
        {
            IntPtr hwndParent = GetParent(hWndChild);

            while (hwndParent != IntPtr.Zero)
            {
                if (GetWindowClassName(hwndParent).Equals(className, StringComparison.InvariantCultureIgnoreCase))
                {
                    return hwndParent;
                }
                hwndParent = GetParent(hwndParent);
            }
            return IntPtr.Zero;
        }
    }
}
