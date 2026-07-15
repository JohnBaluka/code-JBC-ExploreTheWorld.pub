Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_oExcel As Excel.Application
Attribute m_oExcel.VB_VarHelpID = -1
Private m_oWb As Excel.Workbook

' -------------------------------------------------------
' Helper: check if this event should be logged
' -------------------------------------------------------
Private Function ShouldLog(sEventName As String) As Boolean
    ShouldLog = Nz(DLookup("Log", "MsOfficeEvents", "Name = '" & sEventName & "' AND Excel = True"), False)
End Function

' -------------------------------------------------------
' Log area
' -------------------------------------------------------
Private Sub cmd_ClearLog_Click()
    Me.txt_Log.Value = ""
End Sub

' -------------------------------------------------------
' Connect / Disconnect
' The Connect button attaches to a running Excel instance, or
' launches Excel if it is not already open.
' -------------------------------------------------------
Private Sub cmd_Connect_Click()
    If Me.cmd_Connect.Caption = "Connect" Then
        On Error Resume Next

        ' Attach to a running Excel instance...
        Set m_oExcel = GetObject(, "Excel.Application")
        If Err.Number <> 0 Or m_oExcel Is Nothing Then
            ' ...or launch Excel if it is not already open.
            Err.Clear
            Set m_oExcel = New Excel.Application
            If Err.Number <> 0 Or m_oExcel Is Nothing Then
                Err.Clear
                APP.AppendLog Me.txt_Log, "ERROR: Could not start Excel."
                Set m_oExcel = Nothing
                On Error GoTo 0
                Exit Sub
            End If
            APP.AppendLog Me.txt_Log, "Started Excel."
        End If

        m_oExcel.Visible = True
        Set m_oWb = m_oExcel.ActiveWorkbook
        On Error GoTo 0

        Me.cmd_Connect.Caption = "Disconnect"
        If m_oWb Is Nothing Then
            APP.AppendLog Me.txt_Log, "Connected. (No active workbook.)"
        Else
            APP.AppendLog Me.txt_Log, "Connected: " & m_oWb.Name
        End If

        UpdateDefaultJsonPath
        RefreshOpenFilesList
    Else
        Set m_oWb = Nothing
        Set m_oExcel = Nothing
        Me.cmd_Connect.Caption = "Connect"
        APP.AppendLog Me.txt_Log, "Disconnected."
    End If
End Sub

' -------------------------------------------------------
' File selection (output)
' -------------------------------------------------------
Private Sub cmd_SelectFileOutput_Click()
    Dim oDialog As Office.FileDialog
    Set oDialog = Application.FileDialog(msoFileDialogFilePicker)
    With oDialog
        .Title = "Select Output JSON File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "JSON Files", "*.json"
        .Filters.Add "All Files", "*.*"
        If .Show = True Then
            Me.txt_OutputFilePath.Value = .SelectedItems(1)
        End If
    End With
End Sub

' -------------------------------------------------------
' JSON Export
' -------------------------------------------------------
Private Sub cmd_SaveAsJson_Click()
    Me.txt_Log.Value = ""

    If m_oWb Is Nothing Then
        APP.AppendLog Me.txt_Log, "ERROR: No workbook connected. Connect a file first."
        Exit Sub
    End If

    Dim sPath As String
    sPath = Nz(Me.txt_OutputFilePath.Value, "")

    If sPath = "" Then
        APP.AppendLog Me.txt_Log, "ERROR: No output file path specified."
        Exit Sub
    End If

    Dim sExportMethod As String
    sExportMethod = Nz(Me.cmb_ExportMethod.Value, "COM")

    APP.AppendLog Me.txt_Log, "Save as JSON started (" & sExportMethod & "): " & sPath

    Dim dStart As Double
    dStart = Timer

    If sExportMethod = "Direct" Then
        ' Call the macro inside the running Excel add-in directly.
        ' Requires ExploreTheWorld.xlsm to be open and macro-enabled.
        On Error Resume Next
        m_oExcel.Run "WriteActiveWorkbookToJsonFile", sPath
        If Err.Number <> 0 Then
            APP.AppendLog Me.txt_Log, "ERROR (Direct): " & Err.Description & "  -  Ensure ExploreTheWorld.xlsm is open and macros are enabled."
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
        APP.AppendLog Me.txt_Log, "Saved (Direct): " & sPath & "  (" & Format(Timer - dStart, "0.000") & "s)"
    Else
        ' COM: call the module directly from Access (early-bound Excel reference)
        MSO_MsExcel_JsonWriter.WriteWorkbookToJsonFile m_oWb, sPath
        APP.AppendLog Me.txt_Log, "Saved (COM): " & sPath & "  (" & Format(Timer - dStart, "0.000") & "s)"
    End If
End Sub

' -------------------------------------------------------
' Log helper: writes an event line if the event is enabled
' -------------------------------------------------------
Private Sub LogEvt(sName As String, sDetail As String)
    If ShouldLog(sName) Then
        If Len(sDetail) > 0 Then
            APP.AppendLog Me.txt_Log, "[Event] " & sName & ": " & sDetail
        Else
            APP.AppendLog Me.txt_Log, "[Event] " & sName
        End If
    End If
End Sub

' -------------------------------------------------------
' Excel.Application event handlers (full set)
' -------------------------------------------------------
Private Sub m_oExcel_AfterCalculate()
    LogEvt "AfterCalculate", ""
End Sub

Private Sub m_oExcel_NewWorkbook(ByVal Wb As Excel.Workbook)
    LogEvt "NewWorkbook", Wb.Name
End Sub

Private Sub m_oExcel_ProtectedViewWindowActivate(ByVal Pvw As Excel.ProtectedViewWindow)
    LogEvt "ProtectedViewWindowActivate", ""
End Sub

Private Sub m_oExcel_ProtectedViewWindowBeforeClose(ByVal Pvw As Excel.ProtectedViewWindow, ByVal Reason As Excel.XlProtectedViewCloseReason, Cancel As Boolean)
    LogEvt "ProtectedViewWindowBeforeClose", ""
End Sub

Private Sub m_oExcel_ProtectedViewWindowBeforeEdit(ByVal Pvw As Excel.ProtectedViewWindow, Cancel As Boolean)
    LogEvt "ProtectedViewWindowBeforeEdit", ""
End Sub

Private Sub m_oExcel_ProtectedViewWindowDeactivate(ByVal Pvw As Excel.ProtectedViewWindow)
    LogEvt "ProtectedViewWindowDeactivate", ""
End Sub

Private Sub m_oExcel_ProtectedViewWindowOpen(ByVal Pvw As Excel.ProtectedViewWindow)
    LogEvt "ProtectedViewWindowOpen", ""
End Sub

Private Sub m_oExcel_ProtectedViewWindowResize(ByVal Pvw As Excel.ProtectedViewWindow)
    LogEvt "ProtectedViewWindowResize", ""
End Sub

Private Sub m_oExcel_SheetActivate(ByVal Sh As Object)
    LogEvt "SheetActivate", Sh.Name
End Sub

Private Sub m_oExcel_SheetBeforeDelete(ByVal Sh As Object)
    LogEvt "SheetBeforeDelete", Sh.Name
End Sub

Private Sub m_oExcel_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Excel.Range, Cancel As Boolean)
    LogEvt "SheetBeforeDoubleClick", Sh.Name & " / " & Target.Address
End Sub

Private Sub m_oExcel_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Excel.Range, Cancel As Boolean)
    LogEvt "SheetBeforeRightClick", Sh.Name & " / " & Target.Address
End Sub

Private Sub m_oExcel_SheetCalculate(ByVal Sh As Object)
    LogEvt "SheetCalculate", Sh.Name
End Sub

Private Sub m_oExcel_SheetChange(ByVal Sh As Object, ByVal Target As Excel.Range)
    LogEvt "SheetChange", Sh.Name & " / " & Target.Address
End Sub

Private Sub m_oExcel_SheetDeactivate(ByVal Sh As Object)
    LogEvt "SheetDeactivate", Sh.Name
End Sub

Private Sub m_oExcel_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Excel.Hyperlink)
    LogEvt "SheetFollowHyperlink", Sh.Name
End Sub

Private Sub m_oExcel_SheetLensGalleryRenderComplete(ByVal Sh As Object)
    LogEvt "SheetLensGalleryRenderComplete", Sh.Name
End Sub

Private Sub m_oExcel_SheetPivotTableAfterValueChange(ByVal Sh As Object, ByVal TargetPivotTable As Excel.PivotTable, ByVal TargetRange As Excel.Range)
    LogEvt "SheetPivotTableAfterValueChange", Sh.Name
End Sub

Private Sub m_oExcel_SheetPivotTableBeforeAllocateChanges(ByVal Sh As Object, ByVal TargetPivotTable As Excel.PivotTable, ByVal ValueChangeStart As Long, ByVal ValueChangeEnd As Long, Cancel As Boolean)
    LogEvt "SheetPivotTableBeforeAllocateChanges", Sh.Name
End Sub

Private Sub m_oExcel_SheetPivotTableBeforeCommitChanges(ByVal Sh As Object, ByVal TargetPivotTable As Excel.PivotTable, ByVal ValueChangeStart As Long, ByVal ValueChangeEnd As Long, Cancel As Boolean)
    LogEvt "SheetPivotTableBeforeCommitChanges", Sh.Name
End Sub

Private Sub m_oExcel_SheetPivotTableBeforeDiscardChanges(ByVal Sh As Object, ByVal TargetPivotTable As Excel.PivotTable, ByVal ValueChangeStart As Long, ByVal ValueChangeEnd As Long)
    LogEvt "SheetPivotTableBeforeDiscardChanges", Sh.Name
End Sub

Private Sub m_oExcel_SheetPivotTableUpdate(ByVal Sh As Object, ByVal Target As Excel.PivotTable)
    LogEvt "SheetPivotTableUpdate", Sh.Name
End Sub

Private Sub m_oExcel_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)
    LogEvt "SheetSelectionChange", Sh.Name & " / " & Target.Address
End Sub

Private Sub m_oExcel_SheetTableUpdate(ByVal Sh As Object, ByVal Target As Excel.TableObject)
    LogEvt "SheetTableUpdate", Sh.Name
End Sub

Private Sub m_oExcel_WindowActivate(ByVal Wb As Excel.Workbook, ByVal Wn As Excel.Window)
    LogEvt "WindowActivate", Wb.Name
End Sub

Private Sub m_oExcel_WindowDeactivate(ByVal Wb As Excel.Workbook, ByVal Wn As Excel.Window)
    LogEvt "WindowDeactivate", Wb.Name
End Sub

Private Sub m_oExcel_WindowResize(ByVal Wb As Excel.Workbook, ByVal Wn As Excel.Window)
    LogEvt "WindowResize", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookActivate(ByVal Wb As Excel.Workbook)
    Set m_oWb = Wb
    UpdateDefaultJsonPath
    LogEvt "WorkbookActivate", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookAddinInstall(ByVal Wb As Excel.Workbook)
    LogEvt "WorkbookAddinInstall", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookAddinUninstall(ByVal Wb As Excel.Workbook)
    LogEvt "WorkbookAddinUninstall", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookAfterSave(ByVal Wb As Excel.Workbook, ByVal Success As Boolean)
    LogEvt "WorkbookAfterSave", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookAfterXmlExport(ByVal Wb As Excel.Workbook, ByVal Map As Excel.XmlMap, ByVal Url As String, ByVal Result As Excel.XlXmlExportResult)
    LogEvt "WorkbookAfterXmlExport", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookAfterXmlImport(ByVal Wb As Excel.Workbook, ByVal Map As Excel.XmlMap, ByVal IsRefresh As Boolean, ByVal Result As Excel.XlXmlImportResult)
    LogEvt "WorkbookAfterXmlImport", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookBeforeClose(ByVal Wb As Excel.Workbook, Cancel As Boolean)
    LogEvt "WorkbookBeforeClose", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookBeforePrint(ByVal Wb As Excel.Workbook, Cancel As Boolean)
    LogEvt "WorkbookBeforePrint", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookBeforeSave(ByVal Wb As Excel.Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    LogEvt "WorkbookBeforeSave", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookBeforeXmlExport(ByVal Wb As Excel.Workbook, ByVal Map As Excel.XmlMap, ByVal Url As String, Cancel As Boolean)
    LogEvt "WorkbookBeforeXmlExport", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookBeforeXmlImport(ByVal Wb As Excel.Workbook, ByVal Map As Excel.XmlMap, ByVal Url As String, ByVal IsRefresh As Boolean, Cancel As Boolean)
    LogEvt "WorkbookBeforeXmlImport", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookDeactivate(ByVal Wb As Excel.Workbook)
    LogEvt "WorkbookDeactivate", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookModelChange(ByVal Wb As Excel.Workbook, ByVal Changes As Excel.ModelChanges)
    LogEvt "WorkbookModelChange", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookNewChart(ByVal Wb As Excel.Workbook, ByVal Ch As Excel.Chart)
    LogEvt "WorkbookNewChart", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookNewSheet(ByVal Wb As Excel.Workbook, ByVal Sh As Object)
    LogEvt "WorkbookNewSheet", Wb.Name & " / " & Sh.Name
End Sub

Private Sub m_oExcel_WorkbookOpen(ByVal Wb As Excel.Workbook)
    LogEvt "WorkbookOpen", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookPivotTableCloseConnection(ByVal Wb As Excel.Workbook, ByVal Target As Excel.PivotTable)
    LogEvt "WorkbookPivotTableCloseConnection", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookPivotTableOpenConnection(ByVal Wb As Excel.Workbook, ByVal Target As Excel.PivotTable)
    LogEvt "WorkbookPivotTableOpenConnection", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookRowsetComplete(ByVal Wb As Excel.Workbook, ByVal Description As String, ByVal Sheet As String, ByVal Success As Boolean)
    LogEvt "WorkbookRowsetComplete", Wb.Name
End Sub

Private Sub m_oExcel_WorkbookSync(ByVal Wb As Excel.Workbook, ByVal SyncEventType As Office.MsoSyncEventType)
    LogEvt "WorkbookSync", Wb.Name
End Sub

' -------------------------------------------------------
' Active workbook tracking + default JSON path
' -------------------------------------------------------
' Default JSON output follows the standard: the active workbook's
' full path with ".json" appended (e.g. C:\...\filename.xlsx.json).
Private Sub UpdateDefaultJsonPath()
    Dim sFull As String
    On Error Resume Next
    If Not m_oWb Is Nothing Then sFull = m_oWb.FullName
    On Error GoTo 0

    If Len(sFull) > 0 Then
        Me.txt_OutputFilePath.Value = sFull & ".json"
    Else
        Me.txt_OutputFilePath.Value = Environ("USERPROFILE") & "\Documents\ETW_MsExcel.json"
    End If
End Sub

' -------------------------------------------------------
' Open files picklist  -  refreshed every 3 seconds by Timer
' -------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    Me.cmb_ExportMethod.Value = "COM"
    UpdateDefaultJsonPath
    RefreshOpenFilesList
End Sub

Private Sub Form_Timer()
    RefreshOpenFilesList
End Sub

Private Sub RefreshOpenFilesList()
    On Error Resume Next
    Dim oApp As Excel.Application
    Set oApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Or oApp Is Nothing Then
        Err.Clear
        Me.cmb_OpenFiles.RowSource = ""
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' Build list and capture active workbook path
    Dim sList As String
    Dim sActive As String
    Dim oWb As Excel.Workbook
    On Error Resume Next
    sActive = oApp.ActiveWorkbook.FullName
    On Error GoTo 0
    For Each oWb In oApp.Workbooks
        If Len(sList) > 0 Then sList = sList & ";"
        sList = sList & oWb.FullName
    Next
    Set oApp = Nothing
    Me.cmb_OpenFiles.RowSource = sList

    ' Select the active workbook in the picklist
    If Len(sActive) > 0 Then
        Me.cmb_OpenFiles.Value = sActive
    End If
End Sub

Private Sub cmb_OpenFiles_Change()
    Dim sPath As String
    sPath = Nz(Me.cmb_OpenFiles.Value, "")
    If Len(sPath) = 0 Then Exit Sub
    If m_oExcel Is Nothing Then Exit Sub

    ' Activate the chosen workbook so it becomes the connected/active workbook.
    On Error Resume Next
    Dim oWb As Excel.Workbook
    For Each oWb In m_oExcel.Workbooks
        If StrComp(oWb.FullName, sPath, vbTextCompare) = 0 Then
            oWb.Activate
            Set m_oWb = oWb
            Exit For
        End If
    Next
    On Error GoTo 0
    UpdateDefaultJsonPath
End Sub
