using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace ArixcelExplorer.UI.Windows;

internal static class ExplorerWindowFocus
{
    private const int SwRestore = 9;

    public static IntPtr HandleOf(Window window) => new WindowInteropHelper(window).Handle;

    public static void Reclaim(Window window)
    {
        window.Activate();
        var hwnd = HandleOf(window);
        if (hwnd == IntPtr.Zero) return;

        var foreground = GetForegroundWindow();
        var foregroundThread = GetWindowThreadProcessId(foreground, IntPtr.Zero);
        var thisThread = GetCurrentThreadId();
        var attached = false;
        if (foregroundThread != 0 && foregroundThread != thisThread)
        {
            attached = AttachThreadInput(foregroundThread, thisThread, true);
        }

        ShowWindow(hwnd, SwRestore);
        BringWindowToTop(hwnd);
        SetForegroundWindow(hwnd);
        SetFocus(hwnd);

        if (attached)
        {
            AttachThreadInput(foregroundThread, thisThread, false);
        }
    }

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr SetFocus(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr processId);

    [DllImport("user32.dll")]
    private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();
}
