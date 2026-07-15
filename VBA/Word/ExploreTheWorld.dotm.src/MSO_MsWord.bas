Attribute VB_Name = "MSO_MsWord"
Option Explicit

Public Sub WriteActiveDocument()
    Dim sDefaultPath As String
    Dim sOutputFilePath As String

    sDefaultPath = DefaultJsonPath()

    sOutputFilePath = InputBox("Enter output file path:", "Write Document to JSON", sDefaultPath)

    If (sOutputFilePath = "") Then
        Exit Sub
    End If

    Call MSO_MsWord_JsonWriter.WriteDocumentToJsonFile(ActiveDocument, sOutputFilePath)
End Sub

' Default JSON output follows the standard: the active document's full
' path with ".json" appended (e.g. C:\...\filename.docx.json). Falls back
' to the Documents folder for an unsaved document.
Private Function DefaultJsonPath() As String
    Dim sFull As String

    On Error Resume Next
    sFull = ActiveDocument.FullName
    On Error GoTo 0

    If (Len(sFull) > 0 And InStr(sFull, "\") > 0) Then
        DefaultJsonPath = sFull & ".json"
    Else
        DefaultJsonPath = Environ("USERPROFILE") & "\Documents\ETW_MsWord.json"
    End If
End Function

' Application.Run entry point — called from C# Direct path (path-only, uses ActiveDocument)
Public Sub WriteActiveDocumentToJsonFile(sOutputFilePath As String)
    Call MSO_MsWord_JsonWriter.WriteDocumentToJsonFile(ActiveDocument, sOutputFilePath)
End Sub
