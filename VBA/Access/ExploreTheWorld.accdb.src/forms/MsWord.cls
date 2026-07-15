Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_oWord As Word.Application
Attribute m_oWord.VB_VarHelpID = -1
Private m_oDoc As Word.Document

' -------------------------------------------------------
' Helper: check if this event should be logged
' -------------------------------------------------------
Private Function ShouldLog(sEventName As String) As Boolean
    ShouldLog = Nz(DLookup("Log", "MsOfficeEvents", "Name = '" & sEventName & "' AND Word = True"), False)
End Function

' -------------------------------------------------------
' Log area
' -------------------------------------------------------
Private Sub cmd_ClearLog_Click()
    Me.txt_Log.Value = ""
End Sub

' -------------------------------------------------------
' Connect / Disconnect
' The Connect button attaches to a running Word instance, or
' launches Word if it is not already open.
' -------------------------------------------------------
Private Sub cmd_Connect_Click()
    If Me.cmd_Connect.Caption = "Connect" Then
        On Error Resume Next

        ' Attach to a running Word instance...
        Set m_oWord = GetObject(, "Word.Application")
        If Err.Number <> 0 Or m_oWord Is Nothing Then
            ' ...or launch Word if it is not already open.
            Err.Clear
            Set m_oWord = New Word.Application
            If Err.Number <> 0 Or m_oWord Is Nothing Then
                Err.Clear
                APP.AppendLog Me.txt_Log, "ERROR: Could not start Word."
                Set m_oWord = Nothing
                On Error GoTo 0
                Exit Sub
            End If
            APP.AppendLog Me.txt_Log, "Started Word."
        End If

        m_oWord.Visible = True
        Set m_oDoc = m_oWord.ActiveDocument
        On Error GoTo 0

        Me.cmd_Connect.Caption = "Disconnect"
        If m_oDoc Is Nothing Then
            APP.AppendLog Me.txt_Log, "Connected. (No active document.)"
        Else
            APP.AppendLog Me.txt_Log, "Connected: " & m_oDoc.Name
        End If

        UpdateDefaultJsonPath
        RefreshOpenFilesList
    Else
        Set m_oDoc = Nothing
        Set m_oWord = Nothing
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

    If m_oDoc Is Nothing Then
        APP.AppendLog Me.txt_Log, "ERROR: No document connected. Connect a file first."
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
        ' Call the macro inside the running Word add-in directly.
        ' Requires ExploreTheWorld.dotm to be open and macro-enabled.
        On Error Resume Next
        m_oWord.Run "WriteActiveDocumentToJsonFile", sPath
        If Err.Number <> 0 Then
            APP.AppendLog Me.txt_Log, "ERROR (Direct): " & Err.Description & "  -  Ensure ExploreTheWorld.dotm is open and macros are enabled."
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
        APP.AppendLog Me.txt_Log, "Saved (Direct): " & sPath & "  (" & Format(Timer - dStart, "0.000") & "s)"
    Else
        ' COM: call the module directly from Access (early-bound Word reference)
        MSO_MsWord_JsonWriter.WriteDocumentToJsonFile m_oDoc, sPath
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
' Word.Application event handlers (full set)
' -------------------------------------------------------
Private Sub m_oWord_Quit()
    LogEvt "Quit", ""
End Sub

Private Sub m_oWord_NewDocument(ByVal Doc As Word.Document)
    LogEvt "NewDocument", Doc.Name
End Sub

Private Sub m_oWord_DocumentOpen(ByVal Doc As Word.Document)
    LogEvt "DocumentOpen", Doc.Name
End Sub

Private Sub m_oWord_DocumentBeforeClose(ByVal Doc As Word.Document, Cancel As Boolean)
    LogEvt "DocumentBeforeClose", Doc.Name
End Sub

Private Sub m_oWord_DocumentBeforePrint(ByVal Doc As Word.Document, Cancel As Boolean)
    LogEvt "DocumentBeforePrint", Doc.Name
End Sub

Private Sub m_oWord_DocumentBeforeSave(ByVal Doc As Word.Document, SaveAsUI As Boolean, Cancel As Boolean)
    LogEvt "DocumentBeforeSave", Doc.Name
End Sub

Private Sub m_oWord_DocumentChange()
    TrackActiveDocument
    LogEvt "DocumentChange", ActiveDocName()
End Sub

Private Sub m_oWord_DocumentSync(ByVal Doc As Word.Document, ByVal SyncStatus As Office.MsoSyncEventType)
    LogEvt "DocumentSync", Doc.Name
End Sub

Private Sub m_oWord_EPostageInsert(ByVal Doc As Word.Document)
    LogEvt "EPostageInsert", Doc.Name
End Sub

Private Sub m_oWord_EPostageInsertEx(ByVal Doc As Word.Document, ByVal cpFrom As Long, ByVal cpTo As Long, ByVal PostageType As Long, ByVal Recipient As Long, ByVal PostageAmount As Long, ByVal PostageRefundID As Long, ByVal PostageTrackingID As String, ByVal ImageFileName As String, ByVal PromiseMailID As Boolean, Cancel As Boolean)
    LogEvt "EPostageInsertEx", Doc.Name
End Sub

Private Sub m_oWord_EPostagePropertyDialog(ByVal Doc As Word.Document)
    LogEvt "EPostagePropertyDialog", Doc.Name
End Sub

Private Sub m_oWord_MailMergeAfterMerge(ByVal Doc As Word.Document, ByVal DocResult As Word.Document)
    LogEvt "MailMergeAfterMerge", Doc.Name
End Sub

Private Sub m_oWord_MailMergeAfterRecordMerge(ByVal Doc As Word.Document)
    LogEvt "MailMergeAfterRecordMerge", Doc.Name
End Sub

Private Sub m_oWord_MailMergeBeforeMerge(ByVal Doc As Word.Document, ByVal StartRecord As Long, ByVal EndRecord As Long, Cancel As Boolean)
    LogEvt "MailMergeBeforeMerge", Doc.Name
End Sub

Private Sub m_oWord_MailMergeBeforeRecordMerge(ByVal Doc As Word.Document, Cancel As Boolean)
    LogEvt "MailMergeBeforeRecordMerge", Doc.Name
End Sub

Private Sub m_oWord_MailMergeDataSourceLoad(ByVal Doc As Word.Document)
    LogEvt "MailMergeDataSourceLoad", Doc.Name
End Sub

Private Sub m_oWord_MailMergeDataSourceValidate(ByVal Doc As Word.Document, Handled As Boolean)
    LogEvt "MailMergeDataSourceValidate", Doc.Name
End Sub

Private Sub m_oWord_MailMergeDataSourceValidate2(ByVal Doc As Word.Document, Handled As Boolean)
    LogEvt "MailMergeDataSourceValidate2", Doc.Name
End Sub

Private Sub m_oWord_MailMergeWizardSendToCustom(ByVal Doc As Word.Document)
    LogEvt "MailMergeWizardSendToCustom", Doc.Name
End Sub

Private Sub m_oWord_MailMergeWizardStateChange(ByVal Doc As Word.Document, FromState As Long, ToState As Long, Handled As Boolean)
    LogEvt "MailMergeWizardStateChange", Doc.Name
End Sub

Private Sub m_oWord_XMLSelectionChange(ByVal Sel As Word.Selection, ByVal OldXMLNode As Word.XMLNode, ByVal NewXMLNode As Word.XMLNode, Reason As Long)
    LogEvt "XMLSelectionChange", ""
End Sub

Private Sub m_oWord_XMLValidationError(ByVal XMLNode As Word.XMLNode)
    LogEvt "XMLValidationError", ""
End Sub

Private Sub m_oWord_WindowActivate(ByVal Doc As Word.Document, ByVal Wn As Word.Window)
    Set m_oDoc = Doc
    UpdateDefaultJsonPath
    LogEvt "WindowActivate", Doc.Name
End Sub

Private Sub m_oWord_WindowDeactivate(ByVal Doc As Word.Document, ByVal Wn As Word.Window)
    LogEvt "WindowDeactivate", Doc.Name
End Sub

Private Sub m_oWord_WindowSelectionChange(ByVal Sel As Word.Selection)
    LogEvt "WindowSelectionChange", Sel.Text
End Sub

Private Sub m_oWord_WindowBeforeRightClick(ByVal Sel As Word.Selection, Cancel As Boolean)
    LogEvt "WindowBeforeRightClick", ""
End Sub

Private Sub m_oWord_WindowBeforeDoubleClick(ByVal Sel As Word.Selection, Cancel As Boolean)
    LogEvt "WindowBeforeDoubleClick", ""
End Sub

Private Sub m_oWord_WindowSize(ByVal Doc As Word.Document, ByVal Wn As Word.Window)
    LogEvt "WindowSize", Doc.Name
End Sub

Private Sub m_oWord_ProtectedViewWindowOpen(ByVal PvWindow As Word.ProtectedViewWindow)
    LogEvt "ProtectedViewWindowOpen", ""
End Sub

Private Sub m_oWord_ProtectedViewWindowBeforeEdit(ByVal PvWindow As Word.ProtectedViewWindow, Cancel As Boolean)
    LogEvt "ProtectedViewWindowBeforeEdit", ""
End Sub

Private Sub m_oWord_ProtectedViewWindowBeforeClose(ByVal PvWindow As Word.ProtectedViewWindow, ByVal Reason As Long, Cancel As Boolean)
    LogEvt "ProtectedViewWindowBeforeClose", ""
End Sub

Private Sub m_oWord_ProtectedViewWindowSize(ByVal PvWindow As Word.ProtectedViewWindow)
    LogEvt "ProtectedViewWindowSize", ""
End Sub

Private Sub m_oWord_ProtectedViewWindowActivate(ByVal PvWindow As Word.ProtectedViewWindow)
    LogEvt "ProtectedViewWindowActivate", ""
End Sub

Private Sub m_oWord_ProtectedViewWindowDeactivate(ByVal PvWindow As Word.ProtectedViewWindow)
    LogEvt "ProtectedViewWindowDeactivate", ""
End Sub

' -------------------------------------------------------
' Active document tracking + default JSON path
' -------------------------------------------------------
Private Sub TrackActiveDocument()
    On Error Resume Next
    Set m_oDoc = m_oWord.ActiveDocument
    On Error GoTo 0
    UpdateDefaultJsonPath
End Sub

Private Function ActiveDocName() As String
    On Error Resume Next
    If Not m_oDoc Is Nothing Then ActiveDocName = m_oDoc.Name
    On Error GoTo 0
End Function

' Default JSON output follows the standard: the active document's
' full path with ".json" appended (e.g. C:\...\filename.docx.json).
Private Sub UpdateDefaultJsonPath()
    Dim sFull As String
    On Error Resume Next
    If Not m_oDoc Is Nothing Then sFull = m_oDoc.FullName
    On Error GoTo 0

    If Len(sFull) > 0 Then
        Me.txt_OutputFilePath.Value = sFull & ".json"
    Else
        Me.txt_OutputFilePath.Value = Environ("USERPROFILE") & "\Documents\ETW_MsWord.json"
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
    Dim oApp As Word.Application
    Set oApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Or oApp Is Nothing Then
        Err.Clear
        Me.cmb_OpenFiles.RowSource = ""
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' Build list and capture active document path
    Dim sList As String
    Dim sActive As String
    Dim oDoc As Word.Document
    On Error Resume Next
    sActive = oApp.ActiveDocument.FullName
    On Error GoTo 0
    For Each oDoc In oApp.Documents
        If Len(sList) > 0 Then sList = sList & ";"
        sList = sList & oDoc.FullName
    Next
    Set oApp = Nothing
    Me.cmb_OpenFiles.RowSource = sList

    ' Select the active document in the picklist
    If Len(sActive) > 0 Then
        Me.cmb_OpenFiles.Value = sActive
    End If
End Sub

Private Sub cmb_OpenFiles_Change()
    Dim sPath As String
    sPath = Nz(Me.cmb_OpenFiles.Value, "")
    If Len(sPath) = 0 Then Exit Sub
    If m_oWord Is Nothing Then Exit Sub

    ' Activate the chosen document so it becomes the connected/active document.
    On Error Resume Next
    Dim oDoc As Word.Document
    For Each oDoc In m_oWord.Documents
        If StrComp(oDoc.FullName, sPath, vbTextCompare) = 0 Then
            oDoc.Activate
            Set m_oDoc = oDoc
            Exit For
        End If
    Next
    On Error GoTo 0
    UpdateDefaultJsonPath
End Sub
