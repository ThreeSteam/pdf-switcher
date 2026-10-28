$ErrorActionPreference = 'Stop'
Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class WordNativeObject {
    public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll", CharSet=CharSet.Auto)] static extern int GetClassName(IntPtr hwnd, StringBuilder name, int maxCount);
    [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objectId, ref Guid iid, [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object value);
    public static IntPtr FindMainWindow(int processId) {
        IntPtr result = IntPtr.Zero;
        EnumWindows((hwnd, unused) => {
            uint pid; GetWindowThreadProcessId(hwnd, out pid);
            if (pid == (uint)processId) {
                var name = new StringBuilder(64); GetClassName(hwnd, name, name.Capacity);
                if (name.ToString() == "OpusApp") { result = hwnd; return false; }
            }
            return true;
        }, IntPtr.Zero);
        return result;
    }
    public static object FromMainWindow(IntPtr main) {
        object result = null;
        EnumChildWindows(main, (hwnd, unused) => {
            var name = new StringBuilder(64); GetClassName(hwnd, name, name.Capacity);
            if (name.ToString() == "_WwG") {
                Guid childIid = new Guid("00020400-0000-0000-C000-000000000046");
                object childValue = null;
                int childHr = AccessibleObjectFromWindow(hwnd, 0xFFFFFFF0, ref childIid, ref childValue);
                if (childHr == 0 && childValue != null) { result = childValue; return false; }
            }
            return true;
        }, IntPtr.Zero);
        if (result != null) return result;
        Guid iid = new Guid("00020400-0000-0000-C000-000000000046");
        object value = null;
        int hr = AccessibleObjectFromWindow(main, 0xFFFFFFF0, ref iid, ref value);
        if (hr != 0 || value == null) Marshal.ThrowExceptionForHR(hr);
        return value;
    }
}
'@