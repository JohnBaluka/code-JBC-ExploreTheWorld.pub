Attribute VB_Name = "APP"
Option Explicit

Private m_RootPath As String
Private m_ExportPath As String

Function Get_RootPath() As String
    If (m_RootPath = "") Then
        m_RootPath = Nz(DLookup("Value", "APP__Setting", "Name = '" & "RootPath" & "'"), "")
    End If

    Get_RootPath = m_RootPath
End Function

Function Get_ExportPath() As String
    If (m_ExportPath = "") Then
        m_ExportPath = Nz(DLookup("Value", "APP__Setting", "Name = '" & "ExportPath" & "'"), "")
    End If

    Get_ExportPath = m_ExportPath
End Function

Public Sub AppendLog(oLog As TextBox, sMessage As String)
    If oLog Is Nothing Then Exit Sub
    If Len(oLog.Value & "") = 0 Then
        oLog.Value = Format(Now(), "hh:mm:ss") & " " & sMessage
    Else
        oLog.Value = oLog.Value & vbNewLine & Format(Now(), "hh:mm:ss") & " " & sMessage
    End If

    ' Auto-scroll to the newest line as lines are added.
    On Error Resume Next
    oLog.SetFocus
    oLog.SelStart = Len(oLog.Value & "")
    On Error GoTo 0

    DoEvents
End Sub
