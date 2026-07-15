Attribute VB_Name = "MSO_MsPowerPoint"
Option Explicit

Public Sub WriteActivePresentation()
    Dim sDefaultPath As String
    Dim sOutputFilePath As String

    sDefaultPath = DefaultJsonPath()

    sOutputFilePath = InputBox("Enter output file path:", "Write Presentation to JSON", sDefaultPath)

    If (sOutputFilePath = "") Then
        Exit Sub
    End If

    Call MSO_MsPowerPoint_JsonWriter.WritePresentationToJsonFile(ActivePresentation, sOutputFilePath)
End Sub

' Default JSON output follows the standard: the active presentation's full
' path with ".json" appended (e.g. C:\...\filename.pptx.json). Falls back
' to the Documents folder for an unsaved presentation.
Private Function DefaultJsonPath() As String
    Dim sFull As String

    On Error Resume Next
    sFull = ActivePresentation.FullName
    On Error GoTo 0

    If (Len(sFull) > 0 And InStr(sFull, "\") > 0) Then
        DefaultJsonPath = sFull & ".json"
    Else
        DefaultJsonPath = Environ("USERPROFILE") & "\Documents\ETW_MsPowerPoint.json"
    End If
End Function

' Application.Run entry point — called from C# Direct path (path-only, uses ActivePresentation)
Public Sub WriteActivePresentationToJsonFile(sOutputFilePath As String)
    Call MSO_MsPowerPoint_JsonWriter.WritePresentationToJsonFile(ActivePresentation, sOutputFilePath)
End Sub
