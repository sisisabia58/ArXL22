Attribute VB_Name = "ArixcelShortcuts"
' Arixcel Explorer VBA companion — registers global keyboard shortcuts.
' Install this .xlam alongside the ArixcelExplorer VSTO COM add-in.

Private Const COM_ADDIN_PROG_ID As String = "ArixcelExplorer.ComApi"

Public Sub Auto_Open()
    Application.OnKey "^q", "Arixcel_OpenExplorer"
    Application.OnKey "^+q", "Arixcel_OpenDependents"
    Application.OnKey "^+{BS}", "Arixcel_ReturnToOrigin"
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Application.OnKey "^q"
    Application.OnKey "^+q"
    Application.OnKey "^+{BS}"
End Sub

Public Sub Arixcel_OpenExplorer()
    InvokeComAddIn "OpenExplorer"
End Sub

Public Sub Arixcel_OpenDependents()
    InvokeComAddIn "OpenDependents"
End Sub

Public Sub Arixcel_ReturnToOrigin()
    ' Reserved for explorer stack navigation — wired from VSTO coordinator.
End Sub

Private Sub InvokeComAddIn(ByVal methodName As String)
    On Error GoTo Failed
    Dim addIn As COMAddIn
    Set addIn = Application.COMAddIns(COM_ADDIN_PROG_ID)
    If addIn Is Nothing Then
        MsgBox "Arixcel Explorer COM add-in is not loaded.", vbExclamation, "Arixcel"
        Exit Sub
    End If
    Call Application.Run(addIn.ProgId & "!" & methodName)
    Exit Sub
Failed:
    MsgBox "Could not call Arixcel Explorer (" & methodName & "): " & Err.Description, vbCritical, "Arixcel"
End Sub
