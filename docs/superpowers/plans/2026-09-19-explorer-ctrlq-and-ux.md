# Explorer Ctrl+Q, keyboard navigation, and original UI Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make Ctrl+Q open Explorer without an error, keep arrow keys inside Explorer while jumping/highlighting the matching Excel cells, and restyle the window so it reads like original Arixcel (tree, Info/Value/Location, dark-blue selected row, cyan sheet highlight).

**Architecture:** VBA must call the VSTO add-in's COM automation object, not `Application.Run`. Explorer keeps WPF keyboard focus after each navigation; Excel only selects/highlights the range. Values, Info labels, and window chrome are Core/UI changes on top of the existing AST flatten.

**Tech Stack:** C# / .NET Framework 4.8, VSTO Excel add-in, WPF, VBA `.xlam`, xUnit Core tests.

## Global Constraints

- Repo root: `c:\Users\wisnu\Downloads\Arixcel22\ArXL22`
- Do not edit the previous Wave 1 plan file
- Do not implement Formula Map / Compare / utility catalog
- Do not add F2 in-window formula editing or formula-token click-to-select (Wave 3)
- Explorer stays modeless (`Show()`, not `ShowDialog`)
- Keyboard contract stays: Up/Down navigate, Right expand, Left collapse, Ctrl+Q cycle in-window, Enter keep, Esc back
- Precedent sheet highlight default becomes cyan `#7FDBFF` (original “move around” fill), origin stays pink `#F4C2C2`
- Selected Explorer row is dark blue `#0078D7` with white text (original selected row), not pale `#E6F0FA`
- VBA must never use `Application.Run("ArixcelExplorer.ComApi!OpenExplorer")`

---

## Evidence from Excel (why this plan exists)

1. **Ctrl+Q error:** `Could not call Arixcel Explorer (OpenExplorer): Sorry, we couldn't find C:\Users\wisnu\OneDrive\Documents\ArixcelExplorer.ComApi.` Ribbon still opens Explorer. VBA in `src/ArixcelShortcuts/ArixcelShortcuts.bas` calls `Application.Run(addIn.ProgId & "!" & methodName)`, which Excel treats as a workbook path. `ArixcelComApi.OnConnection` does not assign `COMAddIn.Object`. `ThisAddIn` does not override `RequestComAddInAutomationService`.

2. **Arrow keys / cell movement:** `ExcelTraceService.NavigateToAddress` does `sheet.Activate()` + `range.Select()`, which steals focus from the WPF window. After the first Down, arrows move Excel instead of the tree. Range values render as `System.Object[]` because `Value2.ToString()` is used on arrays.

3. **UI vs original screenshot:** Original has Element / Info / Value / Location, tree `+` glyphs, dark-blue selected row, cyan cell fill on the sheet, formula in the lower box, OK. Replica is a white DataGrid with Component/Value/Location, light selection, no Info, no OK.

```text
Original keyboard loop
  arrow in Explorer → select row (dark blue) → Excel shows that range (cyan) → Explorer KEEPS focus

Broken replica loop
  arrow in Explorer → Excel Select() → Excel takes focus → next arrow moves the sheet
```

---

### Task 1: Fix Ctrl+Q COM automation

**Files:**
- Modify: `src/ArixcelExplorer/ThisAddIn.cs`
- Modify: `src/ArixcelExplorer/ComApi/ArixcelComApi.cs`
- Modify: `src/ArixcelShortcuts/ArixcelShortcuts.bas`
- Test: `tests/ArixcelExplorer.Core.Tests/Formulas/ExplorerTreeTests.cs` (no COM test; verification is sideload)

**Interfaces:**
- Consumes: existing `IArixcelComApi` (`OpenExplorer`, `OpenDependents`, `ReturnToOrigin`, `CloseAllExplorers`)
- Produces: VSTO `RequestComAddInAutomationService()` returns `ArixcelComApi`; `COMAddIns("ArixcelExplorer").Object` is that instance; VBA `InvokeComAddIn` calls `.Object` methods

- [ ] **Step 1: Change `ThisAddIn` to expose the automation object**

Replace `src/ArixcelExplorer/ThisAddIn.cs` with:

```csharp
using System;
using ArixcelExplorer.ComApi;
using ArixcelExplorer.Services;

namespace ArixcelExplorer;

public partial class ThisAddIn
{
    private ArixcelComApi? _comApi;

    private void ThisAddIn_Startup(object sender, EventArgs e)
    {
        AddInCoordinator.Initialize(Application);
        _comApi ??= new ArixcelComApi();
    }

    private void ThisAddIn_Shutdown(object sender, EventArgs e)
    {
        AddInCoordinator.CloseAllExplorers();
        AddInCoordinator.ClearFormulaMap();
        _comApi = null;
    }

    protected override object RequestComAddInAutomationService()
    {
        return _comApi ??= new ArixcelComApi();
    }

    private void InternalStartup()
    {
        Startup += ThisAddIn_Startup;
        Shutdown += ThisAddIn_Shutdown;
    }
}
```

- [ ] **Step 2: Assign `COMAddIn.Object` on the companion ComApi add-in**

In `src/ArixcelExplorer/ComApi/ArixcelComApi.cs`, replace empty `OnConnection` with:

```csharp
public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
{
    if (addInInst is Microsoft.Office.Core.COMAddIn comAddIn)
    {
        comAddIn.Object = this;
    }
}
```

Keep the existing `OpenExplorer` / `OpenDependents` / `ReturnToOrigin` method bodies unchanged.

- [ ] **Step 3: Rewrite VBA to call `.Object`, never `Application.Run`**

Replace `InvokeComAddIn` in `src/ArixcelShortcuts/ArixcelShortcuts.bas` with:

```vb
Private Const VSTO_ADDIN_PROG_ID As String = "ArixcelExplorer"
Private Const COM_ADDIN_PROG_ID As String = "ArixcelExplorer.ComApi"

Private Sub InvokeComAddIn(ByVal methodName As String)
    On Error GoTo Failed
    Dim api As Object
    Set api = ResolveApi()
    If api Is Nothing Then
        MsgBox "Arixcel Explorer COM add-in is not loaded.", vbExclamation, "Arixcel"
        Exit Sub
    End If

    Select Case methodName
        Case "OpenExplorer": api.OpenExplorer
        Case "OpenDependents": api.OpenDependents
        Case "ReturnToOrigin": api.ReturnToOrigin
        Case Else
            Err.Raise vbObjectError + 1, "ArixcelShortcuts", "Unknown method " & methodName
    End Select
    Exit Sub
Failed:
    MsgBox "Could not call Arixcel Explorer (" & methodName & "): " & Err.Description, vbCritical, "Arixcel"
End Sub

Private Function ResolveApi() As Object
    On Error Resume Next
    Dim addIn As COMAddIn
    Set addIn = Application.COMAddIns(VSTO_ADDIN_PROG_ID)
    If Not addIn Is Nothing Then
        If addIn.Connect = False Then addIn.Connect = True
        Set ResolveApi = addIn.Object
        If Not ResolveApi Is Nothing Then Exit Function
    End If
    Set addIn = Application.COMAddIns(COM_ADDIN_PROG_ID)
    If Not addIn Is Nothing Then
        If addIn.Connect = False Then addIn.Connect = True
        Set ResolveApi = addIn.Object
    End If
End Function
```

Leave `Auto_Open` OnKey bindings as `^q`, `^+q`, `^{BS}`.

- [ ] **Step 4: Rebuild the `.xlam` so Excel picks up the VBA change**

Run from repo root:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\New-ArixcelShortcutsXlam.ps1
```

Expected: `Saved: ...\dist\ArixcelShortcuts.xlam`

- [ ] **Step 5: Commit**

```powershell
git add src/ArixcelExplorer/ThisAddIn.cs src/ArixcelExplorer/ComApi/ArixcelComApi.cs src/ArixcelShortcuts/ArixcelShortcuts.bas
git commit -m "fix: expose VSTO automation object so Ctrl+Q does not look for a Documents file"
```

---

### Task 2: Format range values instead of `System.Object[]`

**Files:**
- Modify: `src/ArixcelExplorer.Core/Tracing/TraceUtils.cs` (`FormatTraceValue`)
- Modify: `src/ArixcelExplorer/Services/ExcelFormulaService.cs` (`BuildExplorerTree`, `TryResolveReferenceValue`)
- Modify: `src/ArixcelExplorer.Core/Formulas/FormulaEvaluator.cs` (use formatted values already returned as strings)
- Test: `tests/ArixcelExplorer.Core.Tests/Formulas/ExplorerTreeTests.cs`

**Interfaces:**
- Consumes: `object?` from Excel `Value2`
- Produces: `TraceUtils.FormatTraceValue(object? value)` returns a scalar string, or `"n values"` for arrays — never `"System.Object[]"`

- [ ] **Step 1: Write the failing tests**

Add to `tests/ArixcelExplorer.Core.Tests/Formulas/ExplorerTreeTests.cs`:

```csharp
[Fact]
public void FormatTraceValue_does_not_emit_system_object_array()
{
    object[,] block =
    {
        { 10.5, 20d },
        { 30d, 40d }
    };

    Assert.Equal("4 values", TraceUtils.FormatTraceValue(block));
    Assert.Equal("15", TraceUtils.FormatTraceValue(15d));
    Assert.Equal("", TraceUtils.FormatTraceValue(null));
    Assert.Equal("hello", TraceUtils.FormatTraceValue("hello"));
}

[Fact]
public void FormatTraceValue_formats_one_by_n_array_as_joined_preview()
{
    object[] row = { 1d, 2d, 3d };
    Assert.Equal("1, 2, 3", TraceUtils.FormatTraceValue(row));
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
dotnet test tests/ArixcelExplorer.Core.Tests/ArixcelExplorer.Core.Tests.csproj --filter FormatTraceValue
```

Expected: FAIL because `FormatTraceValue` still uses `value.ToString()`.

- [ ] **Step 3: Replace `FormatTraceValue`**

In `src/ArixcelExplorer.Core/Tracing/TraceUtils.cs` replace `FormatTraceValue` with:

```csharp
public static string FormatTraceValue(object? value)
{
    if (value == null || value is DBNull) return "";
    if (value is string s) return s;
    if (value is bool b) return b ? "TRUE" : "FALSE";
    if (value is double or float or decimal or int or long or short or byte)
    {
        return Convert.ToString(value, System.Globalization.CultureInfo.CurrentCulture) ?? "";
    }

    if (value is Array array)
    {
        var flat = Flatten(array);
        if (flat.Count == 0) return "";
        if (flat.Count <= 4) return string.Join(", ", flat);
        return flat.Count + " values";
    }

    var raw = value.ToString() ?? "";
    if (raw.Contains("System.Object")) return "";
    return raw;
}

private static System.Collections.Generic.List<string> Flatten(Array array)
{
    var items = new System.Collections.Generic.List<string>();
    foreach (var item in array)
    {
        if (item is Array nested)
        {
            items.AddRange(Flatten(nested));
        }
        else
        {
            var formatted = FormatTraceValue(item);
            if (formatted.Length > 0) items.Add(formatted);
        }
    }

    return items;
}
```

- [ ] **Step 4: Use the formatter at Excel boundaries**

In `src/ArixcelExplorer/Services/ExcelFormulaService.cs`:

`BuildExplorerTree` value line:

```csharp
var value = TraceUtils.FormatTraceValue(cell.Value2);
```

`TryResolveReferenceValue`:

```csharp
value = TraceUtils.FormatTraceValue(range.Value2);
return true;
```

Add `using ArixcelExplorer.Core.Tracing;` if missing.

- [ ] **Step 5: Run tests to verify they pass**

```powershell
dotnet test tests/ArixcelExplorer.Core.Tests/ArixcelExplorer.Core.Tests.csproj --filter FormatTraceValue
```

Expected: PASS

- [ ] **Step 6: Commit**

```powershell
git add src/ArixcelExplorer.Core/Tracing/TraceUtils.cs src/ArixcelExplorer/Services/ExcelFormulaService.cs tests/ArixcelExplorer.Core.Tests/Formulas/ExplorerTreeTests.cs
git commit -m "fix: format Excel range values instead of System.Object[]"
```

---

### Task 3: Arrow keys keep Explorer focus and cyan-highlight the Excel range

**Files:**
- Modify: `src/ArixcelExplorer.Core/Settings/ArixcelOptions.cs`
- Modify: `src/ArixcelExplorer/Services/ExcelTraceService.cs`
- Modify: `src/ArixcelExplorer/Services/AddInCoordinator.cs`
- Modify: `src/ArixcelExplorer.UI/Windows/ExplorerWindow.xaml.cs`
- Modify: `src/ArixcelExplorer.UI/Windows/DependentsWindow.xaml.cs` (same focus restore)
- Test: `tests/ArixcelExplorer.Core.Tests/Formulas/ExplorerTreeTests.cs` (default color)

**Interfaces:**
- Consumes: `ExplorerViewModel.NavigateRequested`
- Produces: `ExcelTraceService.NavigateToAddress(string address, bool stealFocus)` — when `stealFocus` is false, activate sheet + select range, do not leave Excel as the keyboard target. `ExplorerWindow.RestoreKeyboardFocus()` called after navigation. Default `PrecedentHighlight` is `#7FDBFF`.

- [ ] **Step 1: Write the failing default-color test**

Add:

```csharp
[Fact]
public void Default_precedent_highlight_is_cyan()
{
    Assert.Equal("#7FDBFF", ArixcelOptions.Default.PrecedentHighlight);
}
```

- [ ] **Step 2: Run it**

```powershell
dotnet test tests/ArixcelExplorer.Core.Tests/ArixcelExplorer.Core.Tests.csproj --filter Default_precedent_highlight
```

Expected: FAIL (`PrecedentHighlight` is still `#0563C1`).

- [ ] **Step 3: Change the default**

In `src/ArixcelExplorer.Core/Settings/ArixcelOptions.cs` set:

```csharp
public string PrecedentHighlight { get; set; } = "#7FDBFF";
```

Leave `OriginHighlight` as `#F4C2C2`.

- [ ] **Step 4: Stop Excel from stealing arrow keys**

Replace `NavigateToAddress` in `src/ArixcelExplorer/Services/ExcelTraceService.cs` with:

```csharp
public void NavigateToAddress(string address, bool stealFocus = true)
{
    if (string.IsNullOrWhiteSpace(address)) return;
    var parsed = TraceUtils.ParseWorksheetScopedAddress(address);
    if (parsed == null) return;

    Excel.Worksheet? sheet = null;
    foreach (Excel.Worksheet ws in _app.Worksheets)
    {
        if (string.Equals(ws.Name, parsed.WorksheetName, StringComparison.OrdinalIgnoreCase))
        {
            sheet = ws;
            break;
        }
    }

    if (sheet == null) return;
    var target = sheet.Range[parsed.RangeAddress];
    var previousUpdating = _app.ScreenUpdating;
    try
    {
        _app.ScreenUpdating = false;
        if (!ReferenceEquals(_app.ActiveSheet, sheet))
        {
            sheet.Activate();
        }

        target.Select();
    }
    finally
    {
        _app.ScreenUpdating = previousUpdating;
    }

    if (!stealFocus)
    {
        TryFocusExcelHwndOwner();
    }
}

private void TryFocusExcelHwndOwner()
{
    // No-op here; WPF window restores focus after this returns.
}
```

In `ShowExplorerForActiveCell` navigate handler (`src/ArixcelExplorer/Services/AddInCoordinator.cs`):

```csharp
vm.NavigateRequested += row =>
{
    if (string.IsNullOrWhiteSpace(row.Location)) return;
    _traceService!.NavigateToAddress(row.Location, stealFocus: false);
    _highlightService!.SetTransient(ownerId, row.Location, _options.PrecedentHighlight, origin.OriginAddress);
    window?.Dispatcher.BeginInvoke(new Action(() =>
    {
        window.Activate();
        window.RestoreKeyboardFocus();
    }), System.Windows.Threading.DispatcherPriority.Input);
};
```

Declare `ExplorerWindow? window = null;` before constructing the window, then assign it, so the lambda can call `RestoreKeyboardFocus`.

Add on `ExplorerWindow`:

```csharp
public void RestoreKeyboardFocus()
{
    Activate();
    TreeGrid.Focus();
    if (TreeGrid.SelectedItem != null)
    {
        var row = TreeGrid.ItemContainerGenerator.ContainerFromItem(TreeGrid.SelectedItem) as System.Windows.Controls.DataGridRow;
        row?.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        TreeGrid.Focus();
    }
}
```

In `Window_PreviewKeyDown`, keep handling Up/Down/Left/Right with `e.Handled = true` so DataGrid cannot forward them to Excel.

Apply the same `stealFocus: false` + `RestoreKeyboardFocus` pattern to `DependentsWindow` navigation.

- [ ] **Step 5: Run tests**

```powershell
dotnet test tests/ArixcelExplorer.Core.Tests/ArixcelExplorer.Core.Tests.csproj --filter Default_precedent_highlight
```

Expected: PASS

- [ ] **Step 6: Commit**

```powershell
git add src/ArixcelExplorer.Core/Settings/ArixcelOptions.cs src/ArixcelExplorer/Services/ExcelTraceService.cs src/ArixcelExplorer/Services/AddInCoordinator.cs src/ArixcelExplorer.UI/Windows/ExplorerWindow.xaml.cs src/ArixcelExplorer.UI/Windows/DependentsWindow.xaml.cs tests/ArixcelExplorer.Core.Tests/Formulas/ExplorerTreeTests.cs
git commit -m "fix: keep Explorer focused while arrow keys highlight Excel cells"
```

---

### Task 4: Match original Explorer chrome (Element / Info / dark-blue row / OK)

**Files:**
- Modify: `src/ArixcelExplorer.Core/Formulas/FormulaAstParser.cs` (`FormulaAstNode.Info`)
- Create: `src/ArixcelExplorer.Core/Formulas/FunctionArgInfo.cs`
- Modify: `src/ArixcelExplorer.UI/ViewModels/ExplorerViewModel.cs`
- Modify: `src/ArixcelExplorer.UI/Windows/ExplorerWindow.xaml`
- Modify: `src/ArixcelExplorer.UI/Themes/ArixcelTheme.xaml`
- Test: `tests/ArixcelExplorer.Core.Tests/Formulas/ExplorerTreeTests.cs`

**Interfaces:**
- Consumes: parsed `FormulaAstNode` children
- Produces: `FunctionArgInfo.LabelFor(string functionName, int argumentIndex)` → e.g. `IF` arg 0 = `logical_test`. `ExplorerTreeRow.Info`. Window headers `Element | Info | Value | Location`. Selected row `#0078D7` / white text. Bottom `OK` calls `RequestKeepClose`.

- [ ] **Step 1: Write Info-label tests**

Create `src/ArixcelExplorer.Core/Formulas/FunctionArgInfo.cs` in step 3; first add tests:

```csharp
[Fact]
public void FunctionArgInfo_labels_if_and_sumifs_arguments()
{
    Assert.Equal("logical_test", FunctionArgInfo.LabelFor("IF", 0));
    Assert.Equal("value_if_true", FunctionArgInfo.LabelFor("IF", 1));
    Assert.Equal("value_if_false", FunctionArgInfo.LabelFor("IF", 2));
    Assert.Equal("sum_range", FunctionArgInfo.LabelFor("SUMIFS", 0));
    Assert.Equal("criteria_range1", FunctionArgInfo.LabelFor("SUMIFS", 1));
    Assert.Equal("criteria1", FunctionArgInfo.LabelFor("SUMIFS", 2));
    Assert.Equal("", FunctionArgInfo.LabelFor("SUM", 0));
}
```

- [ ] **Step 2: Run to fail**

```powershell
dotnet test tests/ArixcelExplorer.Core.Tests/ArixcelExplorer.Core.Tests.csproj --filter FunctionArgInfo
```

Expected: FAIL (`FunctionArgInfo` missing).

- [ ] **Step 3: Implement labels and stamp them on AST children**

Create `src/ArixcelExplorer.Core/Formulas/FunctionArgInfo.cs`:

```csharp
using System;

namespace ArixcelExplorer.Core.Formulas;

public static class FunctionArgInfo
{
    public static string LabelFor(string functionName, int argumentIndex)
    {
        var name = (functionName ?? "").ToUpperInvariant();
        return name switch
        {
            "IF" => argumentIndex switch
            {
                0 => "logical_test",
                1 => "value_if_true",
                2 => "value_if_false",
                _ => ""
            },
            "SUMIFS" => argumentIndex == 0
                ? "sum_range"
                : argumentIndex % 2 == 1
                    ? "criteria_range" + ((argumentIndex + 1) / 2)
                    : "criteria" + (argumentIndex / 2),
            "SUMIF" => argumentIndex switch
            {
                0 => "range",
                1 => "criteria",
                2 => "sum_range",
                _ => ""
            },
            "INDEX" => argumentIndex switch
            {
                0 => "array",
                1 => "row_num",
                2 => "column_num",
                _ => ""
            },
            "OFFSET" => argumentIndex switch
            {
                0 => "reference",
                1 => "rows",
                2 => "cols",
                3 => "height",
                4 => "width",
                _ => ""
            },
            _ => ""
        };
    }
}
```

Add `public string Info { get; set; } = "";` to `FormulaAstNode`.

After `node.Children = args.Nodes;` in `TryParseFunction`, add:

```csharp
for (var i = 0; i < node.Children.Count; i++)
{
    node.Children[i].Info = FunctionArgInfo.LabelFor(name, i);
}
```

In `ExplorerViewModel.RefreshRows`, map `Info = node.Info`. Add `public string Info { get; set; } = "";` on `ExplorerTreeRow`.

- [ ] **Step 4: Restyle the window to original chrome**

In `src/ArixcelExplorer.UI/Themes/ArixcelTheme.xaml` replace selection brushes:

```xml
<SolidColorBrush x:Key="ArixcelHighlightBrush" Color="#0078D7" />
<SolidColorBrush x:Key="ArixcelHighlightTextBrush" Color="#FFFFFF" />
<SolidColorBrush x:Key="ArixcelSelectionBrush" Color="#0078D7" />
<SolidColorBrush x:Key="ArixcelWindowBrush" Color="#F0F0F0" />
```

Set Window background to `{StaticResource ArixcelWindowBrush}`.

In `ExplorerWindow.xaml`:
- Remove the hint `TextBlock` from the layout (put the same text on `Window.Title` tooltip only, or a single status line).
- Rename Component header to `Element`.
- Insert Info column (`Width="110"`) between Element and Value.
- DataGrid selected row `#0078D7` / white foreground (update both `SystemColors.HighlightBrushKey` and cell trigger).
- Keep Consolas `+/-` prefix on Element.
- Formula box stays below the grid, no `FORMULA` caption (original has none); `MinHeight="56"`.
- Add a bottom row:

```xml
<DockPanel Grid.Row="3" Margin="0,8,0,0" LastChildFill="False">
  <TextBlock Text="{Binding StatusText}" VerticalAlignment="Center" Foreground="#555" />
  <Button Content="OK" Width="72" Height="24" DockPanel.Dock="Right"
          IsDefault="True" Click="Ok_Click" />
</DockPanel>
```

```csharp
private void Ok_Click(object sender, RoutedEventArgs e) => _viewModel.RequestKeepClose();
```

Strip `'Sheet'!` display noise in the Element column for the origin row only if the label already equals Location; do not change Location binding.

- [ ] **Step 5: Run tests**

```powershell
dotnet test tests/ArixcelExplorer.Core.Tests/ArixcelExplorer.Core.Tests.csproj --filter FunctionArgInfo
```

Expected: PASS

- [ ] **Step 6: Commit**

```powershell
git add src/ArixcelExplorer.Core/Formulas/FunctionArgInfo.cs src/ArixcelExplorer.Core/Formulas/FormulaAstParser.cs src/ArixcelExplorer.UI/ViewModels/ExplorerViewModel.cs src/ArixcelExplorer.UI/Windows/ExplorerWindow.xaml src/ArixcelExplorer.UI/Windows/ExplorerWindow.xaml.cs src/ArixcelExplorer.UI/Themes/ArixcelTheme.xaml tests/ArixcelExplorer.Core.Tests/Formulas/ExplorerTreeTests.cs
git commit -m "feat: restyle Explorer to Element/Info columns and original selection colors"
```

---

### Task 5: Sideload and Excel validation

**Files:**
- Modify: none unless sideload scripts fail
- Test: live Excel

**Interfaces:**
- Consumes: Tasks 1–4 artifacts (`ArixcelExplorer.dll`, `dist/ArixcelShortcuts.xlam`)
- Produces: passing Excel checklist below

- [ ] **Step 1: Close Excel, rebuild, register**

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\Invoke-ArixcelSideload.ps1 -SkipPrerequisites
```

Expected: VSTO + XLAM registered, no Documents path error in output.

- [ ] **Step 2: Enable add-ins if needed**

COM: Arixcel Explorer + Arixcel Explorer API checked. Excel Add-ins: ArixcelShortcuts installed from `dist\ArixcelShortcuts.xlam`.

- [ ] **Step 3: Run the Excel checklist**

Workbook with a nested formula (the SUMIFS on `'Summary (Instit)'!G15` is enough).

1. Select the formula cell. Press **Ctrl+Q**. Explorer opens. **No error dialog.**
2. Window looks gray/classic: Element / Info / Value / Location; selected row dark blue; formula under the grid; OK on the right.
3. Origin cell on the sheet is pink. First tree row is the origin.
4. Press **Down** several times. Explorer selection moves. Excel jumps to that range and fills it **cyan**. Explorer still has focus (next Down still moves the tree, not an unrelated Excel cell).
5. **Right** expands, **Left** collapses.
6. Range rows show `1, 2, 3` or `n values`, never `System.Object[]`.
7. SUMIFS children show Info `sum_range` / `criteria_range1` / `criteria1`.
8. **Enter** or **OK** closes and keeps the current Excel selection.
9. Reopen, **Esc** returns to the origin cell.

- [ ] **Step 4: Commit only if sideload scripts were patched**

If no script changes, skip commit.

---

## Self-review

1. **Spec coverage:** Ctrl+Q error → Task 1. `System.Object[]` → Task 2. Arrow keys + cyan sheet highlight + keep focus → Task 3. Original UI (Element/Info, dark-blue row, OK) → Task 4. Live Excel proof → Task 5.
2. **Placeholders:** none. `stealFocus` and `RestoreKeyboardFocus` are named in Task 3 before later tasks use them.
3. **Type consistency:** `NavigateToAddress(string address, bool stealFocus = true)` is the only navigation signature. `FunctionArgInfo.LabelFor(string, int)` is the only Info API. VBA calls `OpenExplorer` / `OpenDependents` / `ReturnToOrigin` on `IArixcelComApi`.
