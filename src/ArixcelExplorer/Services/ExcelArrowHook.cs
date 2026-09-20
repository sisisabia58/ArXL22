using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using ArixcelExplorer.UI.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArixcelExplorer.Services;

internal sealed class ExcelArrowHook : IDisposable
{
    private const int WhKeyboard = 2;
    private const int HcAction = 0;
    private const int VkLeft = 0x25;
    private const int VkUp = 0x26;
    private const int VkRight = 0x27;
    private const int VkDown = 0x28;
    private const int VkReturn = 0x0D;
    private const int VkEscape = 0x1B;

    private readonly ExplorerSession _session;
    private readonly KeyboardProc _proc;
    private IntPtr _hook;

    public ExcelArrowHook(ExplorerSession session, Excel.Application application)
    {
        _ = application;
        _session = session;
        _proc = Callback;
    }

    public void Install()
    {
        if (_hook != IntPtr.Zero) return;
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
        try
        {
            if (nCode == HcAction && _session.HasOpenWindows && !IsKeyUp(lParam))
            {
                var key = KeyFromVirtualKey((int)wParam);
                if (key != Key.None)
                {
                    var dispatcher = Application.Current?.Dispatcher;
                    if (dispatcher != null)
                    {
                        dispatcher.BeginInvoke(new Action(() => _session.DispatchKey(key)));
                    }
                    else
                    {
                        _session.DispatchKey(key);
                    }

                    return (IntPtr)1;
                }
            }
        }
        catch
        {
            // never throw from a hook
        }

        return CallNextHookEx(_hook, nCode, wParam, lParam);
    }

    private static bool IsKeyUp(IntPtr lParam) => ((uint)lParam.ToInt64() & 0x80000000) != 0;

    private static Key KeyFromVirtualKey(int vk) => vk switch
    {
        VkUp => Key.Up,
        VkDown => Key.Down,
        VkLeft => Key.Left,
        VkRight => Key.Right,
        VkReturn => Key.Enter,
        VkEscape => Key.Escape,
        _ => Key.None
    };

    private delegate IntPtr KeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr SetWindowsHookEx(int idHook, KeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();
}
