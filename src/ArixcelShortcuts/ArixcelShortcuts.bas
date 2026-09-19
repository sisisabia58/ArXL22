Attribute VB_Name = "ArixcelShortcuts"
' Arixcel Explorer VBA companion — registers global keyboard shortcuts.
' Install this .xlam alongside the ArixcelExplorer VSTO COM add-in.

Private Const VSTO_ADDIN_PROG_ID As String = "ArixcelExplorer"
Private Const COM_ADDIN_PROG_ID As String = "ArixcelExplorer.ComApi"

Public Sub Auto_Open()
    Application.OnKey "^q", "Arixcel_OpenExplorer"
    Application.OnKey "^+q", "Arixcel_OpenDependents"
    Application.OnKey "^{BS}", "Arixcel_ReturnToOrigin"
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Application.OnKey "^q"
    Application.OnKey "^+q"
    Application.OnKey "^{BS}"
End Sub

Public Sub Arixcel_OpenExplorer()
    InvokeComAddIn "OpenExplorer"
End Sub

Public Sub Arixcel_OpenDependents()
    InvokeComAddIn "OpenDependents"
End Sub

Public Sub Arixcel_ReturnToOrigin()
    InvokeComAddIn "ReturnToOrigin"
End Sub

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
