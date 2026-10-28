using System;
using System.Runtime.InteropServices;

namespace ArixcelExplorer.UI.Services;

/// <summary>
/// While the formula pane is editing, arrow keys must reach the text box.
/// The installed host hook swallows them for the tree; this hook is installed
/// later on the same thread and lets arrows through without calling that hook.
/// </summary>
internal sealed class FormulaCaretHook : IDisposable
{
    private const int WhKeyboard = 2;
    private readonly NativeProc _proc;
    private IntPtr _hook;

    public FormulaCaretHook()
    {
        _proc = Callback;
        _hook = SetWindowsHookEx(WhKeyboard, _proc, IntPtr.Zero, GetCurrentThreadId());
    }

    public void Dispose()
    {
        if (_hook == IntPtr.Zero) return;
        UnhookWindowsHookEx(_hook);
        _hook = IntPtr.Zero;
    }

    private IntPtr Callback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && IsArrow((int)wParam))
        {
            return IntPtr.Zero;
        }

        return CallNextHookEx(_hook, nCode, wParam, lParam);
    }

    private static bool IsArrow(int vk) => vk is 0x25 or 0x26 or 0x27 or 0x28;

    private delegate IntPtr NativeProc(int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr SetWindowsHookEx(int idHook, NativeProc lpfn, IntPtr hMod, uint threadId);

    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hook);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hook, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();
}
