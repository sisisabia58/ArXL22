using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using ArixcelExplorer.UI.ViewModels;
using ArixcelExplorer.UI.Windows;

namespace ArixcelExplorer.Services;

public sealed class SessionWindow
{
    public SessionWindow(Window window, ExplorerStackEntry origin, string highlightOwnerId)
    {
        Window = window;
        Origin = origin;
        HighlightOwnerId = highlightOwnerId;
    }

    public Window Window { get; }
    public ExplorerStackEntry Origin { get; }
    public string HighlightOwnerId { get; }
}

public sealed class ExplorerSession
{
    private readonly List<SessionWindow> _windows = new();
    private readonly ExplorerStack _origins = new();
    private bool _closingAll;
    private System.Windows.Input.Key _lastKey;
    private int _lastKeyTick;

    public bool HasOpenWindows => _windows.Count > 0;

    public ExplorerStackEntry? FirstOrigin => _origins.First;

    public event Action<SessionWindow, ExplorerCloseMode, ExplorerStackEntry?>? WindowClosed;

    public void Track(Window window, ExplorerStackEntry origin, string highlightOwnerId)
    {
        _origins.Push(origin);
        var sessionWindow = new SessionWindow(window, origin, highlightOwnerId);
        _windows.Add(sessionWindow);
        window.Closed += (_, _) => OnClosed(sessionWindow);
    }

    public ExplorerStackEntry? PreviousOrigin(Window window)
    {
        var index = _windows.FindIndex(item => ReferenceEquals(item.Window, window));
        if (index > 0) return _windows[index - 1].Origin;
        if (index == 0) return _windows[0].Origin;
        return _origins.First;
    }

    public bool DispatchKey(System.Windows.Input.Key key)
    {
        var top = _windows.LastOrDefault()?.Window;
        if (top == null) return false;

        var tick = Environment.TickCount;
        if (key == _lastKey && tick - _lastKeyTick < 20) return true;
        _lastKey = key;
        _lastKeyTick = tick;

        return top switch
        {
            ExplorerWindow explorer => explorer.TryHandleExplorerKey(key),
            DependentsWindow dependents => dependents.TryHandleExplorerKey(key),
            _ => false
        };
    }

    public void ActivateLatest()
    {
        var top = _windows.LastOrDefault();
        if (top == null) return;
        if (top.Window.WindowState == WindowState.Minimized)
        {
            top.Window.WindowState = WindowState.Normal;
        }

        top.Window.Activate();
        top.Window.Focus();
    }

    public void CloseAll()
    {
        _closingAll = true;
        foreach (var item in _windows.ToList())
        {
            try
            {
                item.Window.Close();
            }
            catch
            {
                // ignore close failures during shutdown
            }
        }

        _windows.Clear();
        _origins.Clear();
        _closingAll = false;
    }

    private void OnClosed(SessionWindow sessionWindow)
    {
        if (_closingAll) return;

        var previous = PreviousOrigin(sessionWindow.Window);
        _windows.Remove(sessionWindow);
        var mode = sessionWindow.Window switch
        {
            ExplorerWindow explorer => explorer.CloseMode,
            DependentsWindow dependents => dependents.CloseMode,
            _ => ExplorerCloseMode.KeepSelection
        };
        WindowClosed?.Invoke(sessionWindow, mode, previous);
    }
}
