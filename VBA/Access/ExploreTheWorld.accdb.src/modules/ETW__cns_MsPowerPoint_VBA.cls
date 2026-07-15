VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ETW__cns_MsPowerPoint_VBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' ETW__cns_MsPowerPoint_VBA
' Exports cns_Country data to a PowerPoint presentation
' using the PowerPoint COM object model (Office automation / VBA).
' Title slide ("Explore the World" + country count, no timestamp) followed
' by one centered slide per country (name, large flag, ISO codes) — mirroring
' the "Country Slides" page.
' Class module: m_oPPT is WithEvents so PowerPoint.Application
' events are available for logging and interception.
' ============================================================

Private WithEvents m_oPPT As PowerPoint.Application
Attribute m_oPPT.VB_VarHelpID = -1
Private m_oLog As TextBox

Public Sub Export(oLog As TextBox, sOutputPath As String, sFilter As String, bFilterOn As Boolean)
    Set m_oLog = oLog
    APP.AppendLog m_oLog, "=== CNS Export - PowerPoint (VBA) started ==="

    On Error GoTo ErrHandler

    If Len(Trim(sOutputPath)) = 0 Then
        sOutputPath = APP.Get_ExportPath() & "ETW_cns_MsPowerPoint_VBA.pptx"
    End If

    ' --- Build recordset ---
    Dim db As DAO.Database
    Set db = CurrentDb()
    Dim sSql As String
    sSql = BuildSql(sFilter, bFilterOn)
    APP.AppendLog m_oLog, "  Query: " & sSql
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sSql, dbOpenSnapshot)

    If rs.EOF And rs.BOF Then
        APP.AppendLog m_oLog, "  No records to export."
        rs.Close
        Set db = Nothing
        Set m_oLog = Nothing
        Exit Sub
    End If

    ' --- Open PowerPoint ---
    APP.AppendLog m_oLog, "  Opening PowerPoint..."
    Set m_oPPT = New PowerPoint.Application
    m_oPPT.Visible = msoTrue

    Dim oPres As PowerPoint.Presentation
    Set oPres = m_oPPT.Presentations.Add(msoTrue)

    Dim sngW As Single, sngH As Single, sngMargin As Single
    sngW = oPres.PageSetup.SlideWidth
    sngH = oPres.PageSetup.SlideHeight
    sngMargin = sngW * 0.06

    ' Count the records (for the title slide), then rewind.
    Dim lCount As Long
    rs.MoveLast
    lCount = rs.RecordCount
    rs.MoveFirst

    ' Title slide ("Explore the World" + country count — no timestamp).
    Dim oSlide As PowerPoint.Slide
    Set oSlide = oPres.Slides.Add(1, ppLayoutBlank)
    AddCenteredText oSlide, sngMargin, sngH * 0.32, sngW - 2 * sngMargin, sngH * 0.22, "Explore the World", 40, True
    AddCenteredText oSlide, sngMargin, sngH * 0.56, sngW - 2 * sngMargin, sngH * 0.12, lCount & " countries", 20, False

    Dim lSlide As Long
    Dim sFlagPath As String
    Dim oPicture As PowerPoint.Shape
    lSlide = 2
    Do While Not rs.EOF
        Dim sIso2 As String
        Dim sCountry As String
        Dim sIso3 As String
        Dim sCodes As String
        sIso2 = Nz(rs.Fields("Iso2").Value, "")
        sCountry = Nz(rs.Fields("Country").Value, "")
        sIso3 = Nz(rs.Fields("Iso3").Value, "")

        ' One centered slide per country: name, large flag, ISO codes.
        Set oSlide = oPres.Slides.Add(lSlide, ppLayoutBlank)
        AddCenteredText oSlide, sngMargin, sngH * 0.06, sngW - 2 * sngMargin, sngH * 0.2, sCountry, 36, True

        sFlagPath = ETW__FlagImages.Get_FlagImageFilePath( _
            m_oLog, sIso2, Nz(rs.Fields("FlagUrl").Value, ""))
        If Len(sFlagPath) > 0 Then
            ' Insert at original size, then scale by height with the aspect ratio locked and center it.
            Set oPicture = oSlide.Shapes.AddPicture(sFlagPath, msoFalse, msoCTrue, _
                0, sngH * 0.3, -1, -1)
            oPicture.LockAspectRatio = msoTrue
            oPicture.Height = sngH * 0.42
            oPicture.Left = (sngW - oPicture.Width) / 2
        End If

        sCodes = sIso2
        If Len(sIso3) > 0 Then sCodes = sCodes & "      " & sIso3
        AddCenteredText oSlide, sngMargin, sngH * 0.8, sngW - 2 * sngMargin, sngH * 0.12, sCodes, 20, False

        lSlide = lSlide + 1
        rs.MoveNext
    Loop

    rs.Close
    Set db = Nothing

    APP.AppendLog m_oLog, "  Saving to: " & sOutputPath
    oPres.SaveAs sOutputPath, ppSaveAsOpenXMLPresentation
    oPres.Close
    m_oPPT.Quit
    Set oSlide = Nothing
    Set oPres = Nothing
    Set m_oPPT = Nothing

    APP.AppendLog m_oLog, "=== CNS Export - PowerPoint (VBA) complete ==="
    Set m_oLog = Nothing
    Exit Sub

ErrHandler:
    APP.AppendLog m_oLog, "ERROR: " & Err.Description
    On Error Resume Next
    If Not oPres Is Nothing Then oPres.Close
    If Not m_oPPT Is Nothing Then m_oPPT.Quit
    Set m_oPPT = Nothing
    Set m_oLog = Nothing
    On Error GoTo 0
End Sub

' Adds a horizontally- and vertically-centered text box to a slide.
Private Sub AddCenteredText(oSlide As PowerPoint.Slide, sngLeft As Single, sngTop As Single, _
        sngWidth As Single, sngHeight As Single, sText As String, sngSize As Single, bBold As Boolean)
    Dim oBox As PowerPoint.Shape
    Set oBox = oSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, sngLeft, sngTop, sngWidth, sngHeight)
    With oBox.TextFrame
        .TextRange.Text = sText
        .TextRange.Font.Size = sngSize
        If bBold Then
            .TextRange.Font.Bold = msoTrue
        Else
            .TextRange.Font.Bold = msoFalse
        End If
        .TextRange.ParagraphFormat.Alignment = ppAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoTrue
    End With
End Sub

Private Function BuildSql(sFilter As String, bFilterOn As Boolean) As String
    Dim sSql As String
    sSql = "SELECT cns_Country.*, cns_CountryFlag.Flag AS FlagUrl FROM cns_Country" & _
           " LEFT JOIN cns_CountryFlag ON cns_Country.Iso2 = cns_CountryFlag.Iso2"
    If bFilterOn And Len(Trim(sFilter)) > 0 Then
        sSql = sSql & " WHERE " & sFilter
    End If
    sSql = sSql & " ORDER BY cns_Country.Iso2"
    BuildSql = sSql
End Function

' -------------------------------------------------------
' PowerPoint.Application events
' -------------------------------------------------------
Private Sub m_oPPT_AfterNewPresentation(ByVal Pres As PowerPoint.Presentation)
    APP.AppendLog m_oLog, "  [Event] AfterNewPresentation: " & Pres.Name
End Sub

Private Sub m_oPPT_PresentationBeforeSave(ByVal Pres As PowerPoint.Presentation, Cancel As Boolean)
    APP.AppendLog m_oLog, "  [Event] PresentationBeforeSave: " & Pres.Name
End Sub

Private Sub m_oPPT_PresentationBeforeClose(ByVal Pres As PowerPoint.Presentation, Cancel As Boolean)
    APP.AppendLog m_oLog, "  [Event] PresentationBeforeClose: " & Pres.Name
End Sub
