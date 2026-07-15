Attribute VB_Name = "MSO_MsExcel"
Option Explicit

Public Sub WriteActiveWorkbook()
    Dim sDefaultPath As String
    Dim sOutputFilePath As String

    sDefaultPath = DefaultJsonPath()

    sOutputFilePath = InputBox("Enter output file path:", "Write Workbook to JSON", sDefaultPath)

    If (sOutputFilePath = "") Then
        Exit Sub
    End If

    Call MSO_MsExcel_JsonWriter.WriteWorkbookToJsonFile(ActiveWorkbook, sOutputFilePath)
End Sub

' Default JSON output follows the standard: the active workbook's full
' path with ".json" appended (e.g. C:\...\filename.xlsx.json). Falls back
' to the Documents folder for an unsaved workbook.
Private Function DefaultJsonPath() As String
    Dim sFull As String

    On Error Resume Next
    sFull = ActiveWorkbook.FullName
    On Error GoTo 0

    If (Len(sFull) > 0 And InStr(sFull, "\") > 0) Then
        DefaultJsonPath = sFull & ".json"
    Else
        DefaultJsonPath = Environ("USERPROFILE") & "\Documents\ETW_MsExcel.json"
    End If
End Function

' Application.Run entry point — called from C# Direct path (path-only, uses ActiveWorkbook)
Public Sub WriteActiveWorkbookToJsonFile(sOutputFilePath As String)
    Call MSO_MsExcel_JsonWriter.WriteWorkbookToJsonFile(ActiveWorkbook, sOutputFilePath)
End Sub
