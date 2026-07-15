Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_oPPT As PowerPoint.Application
Attribute m_oPPT.VB_VarHelpID = -1
Private m_oPres As PowerPoint.Presentation

' -------------------------------------------------------
' Helper: check if this event should be logged
' -------------------------------------------------------
Private Function ShouldLog(sEventName As String) As Boolean
    ShouldLog = Nz(DLookup("Log", "MsOfficeEvents", "Name = '" & sEventName & "' AND PowerPoint = True"), False)
End Function

' -------------------------------------------------------
' Log area
' -------------------------------------------------------
Private Sub cmd_ClearLog_Click()
    Me.txt_Log.Value = ""
End Sub

' -------------------------------------------------------
' Connect / Disconnect
' The Connect button attaches to a running PowerPoint instance,
' or launches PowerPoint if it is not already open.
' -------------------------------------------------------
Private Sub cmd_Connect_Click()
    If Me.cmd_Connect.Caption = "Connect" Then
        On Error Resume Next

        ' Attach to a running PowerPoint instance...
        Set m_oPPT = GetObject(, "PowerPoint.Application")
        If Err.Number <> 0 Or m_oPPT Is Nothing Then
            ' ...or launch PowerPoint if it is not already open.
            Err.Clear
            Set m_oPPT = New PowerPoint.Application
            If Err.Number <> 0 Or m_oPPT Is Nothing Then
                Err.Clear
                APP.AppendLog Me.txt_Log, "ERROR: Could not start PowerPoint."
                Set m_oPPT = Nothing
                On Error GoTo 0
                Exit Sub
            End If
            APP.AppendLog Me.txt_Log, "Started PowerPoint."
        End If

        m_oPPT.Visible = True
        Set m_oPres = m_oPPT.ActivePresentation
        On Error GoTo 0

        Me.cmd_Connect.Caption = "Disconnect"
        If m_oPres Is Nothing Then
            APP.AppendLog Me.txt_Log, "Connected. (No active presentation.)"
        Else
            APP.AppendLog Me.txt_Log, "Connected: " & m_oPres.Name
        End If

        UpdateDefaultJsonPath
        RefreshOpenFilesList
    Else
        Set m_oPres = Nothing
        Set m_oPPT = Nothing
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

    If m_oPres Is Nothing Then
        APP.AppendLog Me.txt_Log, "ERROR: No presentation connected. Connect a file first."
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
        ' Call the macro inside the running PowerPoint add-in directly.
        ' Requires ExploreTheWorld.pptm to be open and macro-enabled.
        On Error Resume Next
        m_oPPT.Run "WriteActivePresentationToJsonFile", sPath
        If Err.Number <> 0 Then
            APP.AppendLog Me.txt_Log, "ERROR (Direct): " & Err.Description & "  -  Ensure ExploreTheWorld.pptm is open and macros are enabled."
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
        APP.AppendLog Me.txt_Log, "Saved (Direct): " & sPath & "  (" & Format(Timer - dStart, "0.000") & "s)"
    Else
        ' COM: call the module directly from Access (early-bound PowerPoint reference)
        MSO_MsPowerPoint_JsonWriter.WritePresentationToJsonFile m_oPres, sPath
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
' PowerPoint.Application event handlers (full set)
' -------------------------------------------------------
Private Sub m_oPPT_AfterNewPresentation(ByVal Pres As PowerPoint.Presentation)
    LogEvt "AfterNewPresentation", Pres.Name
End Sub

Private Sub m_oPPT_AfterPresentationOpen(ByVal Pres As PowerPoint.Presentation)
    LogEvt "AfterPresentationOpen", Pres.Name
End Sub

Private Sub m_oPPT_ColorSchemeChanged(ByVal SldRange As PowerPoint.SlideRange)
    LogEvt "ColorSchemeChanged", ""
End Sub

Private Sub m_oPPT_NewPresentation(ByVal Pres As PowerPoint.Presentation)
    LogEvt "NewPresentation", Pres.Name
End Sub

Private Sub m_oPPT_PresentationBeforeClose(ByVal Pres As PowerPoint.Presentation, Cancel As Boolean)
    LogEvt "PresentationBeforeClose", Pres.Name
End Sub

Private Sub m_oPPT_PresentationBeforeSave(ByVal Pres As PowerPoint.Presentation, Cancel As Boolean)
    LogEvt "PresentationBeforeSave", Pres.Name
End Sub

Private Sub m_oPPT_PresentationClose(ByVal Pres As PowerPoint.Presentation)
    LogEvt "PresentationClose", Pres.Name
End Sub

Private Sub m_oPPT_PresentationCloseFinal(ByVal Pres As PowerPoint.Presentation)
    LogEvt "PresentationCloseFinal", Pres.Name
End Sub

Private Sub m_oPPT_PresentationOpen(ByVal Pres As PowerPoint.Presentation)
    LogEvt "PresentationOpen", Pres.Name
End Sub

Private Sub m_oPPT_PresentationPrint(ByVal Pres As PowerPoint.Presentation)
    LogEvt "PresentationPrint", Pres.Name
End Sub

Private Sub m_oPPT_PresentationSave(ByVal Pres As PowerPoint.Presentation)
    LogEvt "PresentationSave", Pres.Name
End Sub

Private Sub m_oPPT_PresentationSync(ByVal Pres As PowerPoint.Presentation, ByVal SyncEventType As Office.MsoSyncEventType)
    LogEvt "PresentationSync", Pres.Name
End Sub

Private Sub m_oPPT_AfterShapeSizeChange(ByVal shp As PowerPoint.Shape)
    LogEvt "AfterShapeSizeChange", ""
End Sub

Private Sub m_oPPT_AfterDragDropOnSlide(ByVal Sld As PowerPoint.Slide, ByVal X As Single, ByVal Y As Single)
    LogEvt "AfterDragDropOnSlide", ""
End Sub

Private Sub m_oPPT_PresentationNewSlide(ByVal Sld As PowerPoint.Slide)
    LogEvt "PresentationNewSlide", ""
End Sub

Private Sub m_oPPT_SlideSelectionChanged(ByVal SldRange As PowerPoint.SlideRange)
    LogEvt "SlideSelectionChanged", ""
End Sub

Private Sub m_oPPT_SlideShowBegin(ByVal Wn As PowerPoint.SlideShowWindow)
    LogEvt "SlideShowBegin", Wn.Presentation.Name
End Sub

Private Sub m_oPPT_SlideShowEnd(ByVal Pres As PowerPoint.Presentation)
    LogEvt "SlideShowEnd", Pres.Name
End Sub

Private Sub m_oPPT_SlideShowNextBuild(ByVal Wn As PowerPoint.SlideShowWindow)
    LogEvt "SlideShowNextBuild", ""
End Sub

Private Sub m_oPPT_SlideShowNextClick(ByVal Wn As PowerPoint.SlideShowWindow, ByVal nEffect As PowerPoint.Effect)
    LogEvt "SlideShowNextClick", ""
End Sub

Private Sub m_oPPT_SlideShowNextSlide(ByVal Wn As PowerPoint.SlideShowWindow)
    LogEvt "SlideShowNextSlide", "Slide " & Wn.View.CurrentShowPosition
End Sub

Private Sub m_oPPT_SlideShowOnNext(ByVal Wn As PowerPoint.SlideShowWindow)
    LogEvt "SlideShowOnNext", ""
End Sub

Private Sub m_oPPT_SlideShowOnPrevious(ByVal Wn As PowerPoint.SlideShowWindow)
    LogEvt "SlideShowOnPrevious", ""
End Sub

Private Sub m_oPPT_WindowActivate(ByVal Pres As PowerPoint.Presentation, ByVal Wn As PowerPoint.DocumentWindow)
    Set m_oPres = Pres
    UpdateDefaultJsonPath
    LogEvt "WindowActivate", Pres.Name
End Sub

Private Sub m_oPPT_WindowDeactivate(ByVal Pres As PowerPoint.Presentation, ByVal Wn As PowerPoint.DocumentWindow)
    LogEvt "WindowDeactivate", Pres.Name
End Sub

Private Sub m_oPPT_WindowBeforeRightClick(ByVal Sel As PowerPoint.Selection, Cancel As Boolean)
    LogEvt "WindowBeforeRightClick", ""
End Sub

Private Sub m_oPPT_WindowBeforeDoubleClick(ByVal Sel As PowerPoint.Selection, Cancel As Boolean)
    LogEvt "WindowBeforeDoubleClick", ""
End Sub

Private Sub m_oPPT_WindowSelectionChange(ByVal Sel As PowerPoint.Selection)
    LogEvt "WindowSelectionChange", SelectionInfo(Sel)
End Sub

Private Sub m_oPPT_ProtectedViewWindowOpen(ByVal ProtViewWindow As PowerPoint.ProtectedViewWindow)
    LogEvt "ProtectedViewWindowOpen", ""
End Sub

Private Sub m_oPPT_ProtectedViewWindowBeforeEdit(ByVal ProtViewWindow As PowerPoint.ProtectedViewWindow, Cancel As Boolean)
    LogEvt "ProtectedViewWindowBeforeEdit", ""
End Sub

Private Sub m_oPPT_ProtectedViewWindowBeforeClose(ByVal ProtViewWindow As PowerPoint.ProtectedViewWindow, ByVal Reason As PowerPoint.PpProtectedViewCloseReason, Cancel As Boolean)
    LogEvt "ProtectedViewWindowBeforeClose", ""
End Sub

Private Sub m_oPPT_ProtectedViewWindowActivate(ByVal ProtViewWindow As PowerPoint.ProtectedViewWindow)
    LogEvt "ProtectedViewWindowActivate", ""
End Sub

Private Sub m_oPPT_ProtectedViewWindowDeactivate(ByVal ProtViewWindow As PowerPoint.ProtectedViewWindow)
    LogEvt "ProtectedViewWindowDeactivate", ""
End Sub

' Describe the current PowerPoint selection for the event log.
Private Function SelectionInfo(ByVal Sel As PowerPoint.Selection) As String
    On Error Resume Next
    Select Case Sel.Type
        Case ppSelectionText
            SelectionInfo = Sel.TextRange.Text
        Case ppSelectionShapes
            If Sel.ShapeRange.Count > 0 Then SelectionInfo = Sel.ShapeRange(1).Name
        Case ppSelectionSlides
            If Sel.SlideRange.Count > 0 Then SelectionInfo = "Slide " & Sel.SlideRange(1).SlideIndex
        Case Else
            SelectionInfo = "(none)"
    End Select
    On Error GoTo 0
End Function

' -------------------------------------------------------
' Active presentation tracking + default JSON path
' -------------------------------------------------------
' Default JSON output follows the standard: the active presentation's
' full path with ".json" appended (e.g. C:\...\filename.pptx.json).
Private Sub UpdateDefaultJsonPath()
    Dim sFull As String
    On Error Resume Next
    If Not m_oPres Is Nothing Then sFull = m_oPres.FullName
    On Error GoTo 0

    If Len(sFull) > 0 Then
        Me.txt_OutputFilePath.Value = sFull & ".json"
    Else
        Me.txt_OutputFilePath.Value = Environ("USERPROFILE") & "\Documents\ETW_MsPowerPoint.json"
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
    Dim oApp As PowerPoint.Application
    Set oApp = GetObject(, "PowerPoint.Application")
    If Err.Number <> 0 Or oApp Is Nothing Then
        Err.Clear
        Me.cmb_OpenFiles.RowSource = ""
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' Build list and capture active presentation path
    Dim sList As String
    Dim sActive As String
    Dim oPres As PowerPoint.Presentation
    On Error Resume Next
    sActive = oApp.ActivePresentation.FullName
    On Error GoTo 0
    For Each oPres In oApp.Presentations
        If Len(sList) > 0 Then sList = sList & ";"
        sList = sList & oPres.FullName
    Next
    Set oApp = Nothing
    Me.cmb_OpenFiles.RowSource = sList

    ' Select the active presentation in the picklist
    If Len(sActive) > 0 Then
        Me.cmb_OpenFiles.Value = sActive
    End If
End Sub

Private Sub cmb_OpenFiles_Change()
    Dim sPath As String
    sPath = Nz(Me.cmb_OpenFiles.Value, "")
    If Len(sPath) = 0 Then Exit Sub
    If m_oPPT Is Nothing Then Exit Sub

    ' Activate the chosen presentation so it becomes the connected/active one.
    On Error Resume Next
    Dim oPres As PowerPoint.Presentation
    For Each oPres In m_oPPT.Presentations
        If StrComp(oPres.FullName, sPath, vbTextCompare) = 0 Then
            If oPres.Windows.Count > 0 Then oPres.Windows(1).Activate
            Set m_oPres = oPres
            Exit For
        End If
    Next
    On Error GoTo 0
    UpdateDefaultJsonPath
End Sub
