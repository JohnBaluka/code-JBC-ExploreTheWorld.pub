Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' -- Helpers ------------------------------------------------------------------

Private Function DefaultExportFileName() As String
    Dim sType As String
    Dim sExt  As String
    sType = Nz(Me.cmb_ExportType.Value, "Word")
    Select Case sType
        Case "Excel":       sExt = "xlsx"
        Case "PowerPoint":  sExt = "pptx"
        Case Else:          sExt = "docx"
    End Select
    DefaultExportFileName = Environ("USERPROFILE") & "\Documents\ETW_CountriesNow." & sExt
End Function

Private Sub UpdateDefaultExportPath()
    If Nz(Me.txt_ExportFilePath.Value, "") = "" Then
        Me.txt_ExportFilePath.Value = DefaultExportFileName()
    End If
End Sub

' -- Form events ---------------------------------------------------------------

Private Sub Form_Open(Cancel As Integer)
    Me.cmb_ExportType.Value = "Word"
    Me.txt_ExportFilePath.Value = DefaultExportFileName()
End Sub

' -- Load / Clear --------------------------------------------------------------

Private Sub cmd_Load_Click()
    Me.txt_Log.Value = ""
    ETW__cns_API.Load Me.txt_Log
End Sub

Private Sub cmd_Clear_Click()
    Me.txt_Log.Value = ""
    ETW__cns_API.Clear Me.txt_Log
End Sub

' -- Export Type after update — refresh default file name extension ------------------
Private Sub cmb_ExportType_AfterUpdate()
    ' Regenerate the default path when type changes (only if still a default-looking name)
    Dim sCurrent As String
    sCurrent = Nz(Me.txt_ExportFilePath.Value, "")
    If InStr(sCurrent, "ETW_CountriesNow") > 0 Then
        Me.txt_ExportFilePath.Value = DefaultExportFileName()
    End If
End Sub

' -- Log -----------------------------------------------------------------------

Private Sub cmd_ClearLog_Click()
    Me.txt_Log.Value = ""
End Sub

' -- File picker ---------------------------------------------------------------

Private Sub cmd_SelectExportFile_Click()
    Dim oDialog As Office.FileDialog
    Dim sFilter As String

    Select Case Nz(Me.cmb_ExportType.Value, "")
        Case "Word":        sFilter = "Word Documents (*.docx)|*.docx"
        Case "Excel":       sFilter = "Excel Workbooks (*.xlsx)|*.xlsx"
        Case "PowerPoint":  sFilter = "PowerPoint Presentations (*.pptx)|*.pptx"
        Case Else:          sFilter = "All Files (*.*)|*.*"
    End Select

    Set oDialog = Application.FileDialog(msoFileDialogSaveAs)
    With oDialog
        .Title = "Select Export File Location"
        .InitialFileName = Nz(Me.txt_ExportFilePath.Value, DefaultExportFileName())
        If .Show = True Then
            Me.txt_ExportFilePath.Value = .SelectedItems(1)
        End If
    End With
    Set oDialog = Nothing
End Sub

' -- Export --------------------------------------------------------------------

Private Sub cmd_Export_Click()
    Me.txt_Log.Value = ""

    Dim sType As String
    Dim sFilePath As String
    sType = Nz(Me.cmb_ExportType.Value, "")
    sFilePath = Nz(Me.txt_ExportFilePath.Value, "")

    If sType = "" Then
        APP.AppendLog Me.txt_Log, "ERROR: Please select an Export Type."
        Exit Sub
    End If

    If sFilePath = "" Then
        sFilePath = DefaultExportFileName()
        Me.txt_ExportFilePath.Value = sFilePath
    End If

    Dim dStart As Double
    dStart = Timer

    Select Case sType
        Case "Word"
            ETW__cns_MsWord_VBA.Export Me.txt_Log, sFilePath, Me.Filter, Me.FilterOn
        Case "Excel"
            ETW__cns_MsExcel_VBA.Export Me.txt_Log, sFilePath, Me.Filter, Me.FilterOn
        Case "PowerPoint"
            ETW__cns_MsPowerPoint_VBA.Export Me.txt_Log, sFilePath, Me.Filter, Me.FilterOn
        Case Else
            APP.AppendLog Me.txt_Log, "ERROR: Unknown export type: " & sType
            Exit Sub
    End Select

    ' Note completion + elapsed time. (The Watcher form is intentionally
    ' not opened after an export.)
    APP.AppendLog Me.txt_Log, "Export finished: " & sFilePath & "  (" & Format(Timer - dStart, "0.000") & "s)"
End Sub
