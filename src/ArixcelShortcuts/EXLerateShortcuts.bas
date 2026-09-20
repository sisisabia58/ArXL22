Attribute VB_Name = "EXLerateShortcuts"
' EXLerate Explorer VBA companion — registers global keyboard shortcuts.
' Install this .xlam alongside the EXLerate Explorer VSTO COM add-in.

Private Const VSTO_ADDIN_PROG_ID As String = "EXLerateExplorer"
Private Const COM_ADDIN_PROG_ID As String = "EXLerateExplorer.ComApi"
Private Const LEGACY_VSTO_ADDIN_PROG_ID As String = "ArixcelExplorer"
Private Const LEGACY_COM_ADDIN_PROG_ID As String = "ArixcelExplorer.ComApi"

Public Sub Auto_Open()
    Application.OnKey "^q", "EXLerate_OpenExplorer"
    Application.OnKey "^+q", "EXLerate_OpenDependents"
    Application.OnKey "^{BS}", "EXLerate_ReturnToOrigin"
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Application.OnKey "^q"
    Application.OnKey "^+q"
    Application.OnKey "^{BS}"
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "{RETURN}"
    Application.OnKey "{ESC}"
End Sub

Public Sub EXLerate_OpenExplorer()
    InvokeComAddIn "OpenExplorer"
End Sub

Public Sub EXLerate_OpenDependents()
    InvokeComAddIn "OpenDependents"
End Sub

Public Sub EXLerate_ReturnToOrigin()
    InvokeComAddIn "ReturnToOrigin"
End Sub

Public Sub EXLerate_ExplorerKeyUp()
    InvokeExplorerKey "Up"
End Sub

Public Sub EXLerate_ExplorerKeyDown()
    InvokeExplorerKey "Down"
End Sub

Public Sub EXLerate_ExplorerKeyLeft()
    InvokeExplorerKey "Left"
End Sub

Public Sub EXLerate_ExplorerKeyRight()
    InvokeExplorerKey "Right"
End Sub

Public Sub EXLerate_ExplorerKeyEnter()
    InvokeExplorerKey "Enter"
End Sub

Public Sub EXLerate_ExplorerKeyEscape()
    InvokeExplorerKey "Escape"
End Sub

Private Sub InvokeExplorerKey(ByVal keyName As String)
    On Error Resume Next
    Dim api As Object
    Set api = ResolveApi()
    If api Is Nothing Then Exit Sub
    api.DispatchExplorerKey keyName
End Sub

Private Sub InvokeComAddIn(ByVal methodName As String)
    On Error GoTo Failed
    Dim api As Object
    Set api = ResolveApi()
    If api Is Nothing Then
        MsgBox "EXLerate Explorer COM add-in is not loaded.", vbExclamation, "EXLerate"
        Exit Sub
    End If

    Select Case methodName
        Case "OpenExplorer": api.OpenExplorer
        Case "OpenDependents": api.OpenDependents
        Case "ReturnToOrigin": api.ReturnToOrigin
        Case Else
            Err.Raise vbObjectError + 1, "EXLerateShortcuts", "Unknown method " & methodName
    End Select
    Exit Sub
Failed:
    MsgBox "Could not call EXLerate Explorer (" & methodName & "): " & Err.Description, vbCritical, "EXLerate"
End Sub

Private Function ResolveApi() As Object
    On Error Resume Next
    ' ComApi OnConnection always initializes. VSTO .Object can exist before Startup.
    Set ResolveApi = TryResolve(COM_ADDIN_PROG_ID)
    If Not ResolveApi Is Nothing Then Exit Function
    Set ResolveApi = TryResolve(LEGACY_COM_ADDIN_PROG_ID)
    If Not ResolveApi Is Nothing Then Exit Function
    Set ResolveApi = TryResolve(VSTO_ADDIN_PROG_ID)
    If Not ResolveApi Is Nothing Then Exit Function
    Set ResolveApi = TryResolve(LEGACY_VSTO_ADDIN_PROG_ID)
End Function

Private Function TryResolve(ByVal progId As String) As Object
    On Error Resume Next
    Dim addIn As COMAddIn
    Set addIn = Application.COMAddIns(progId)
    If Not addIn Is Nothing Then
        If addIn.Connect = False Then addIn.Connect = True
        Set TryResolve = addIn.Object
    End If
End Function
