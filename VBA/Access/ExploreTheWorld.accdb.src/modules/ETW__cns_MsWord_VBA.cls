VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ETW__cns_MsWord_VBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' ETW__cns_MsWord_VBA
' Exports cns_Country data to a Microsoft Word document
' using the Word COM object model (Office automation / VBA).
' Class module: m_oWord is WithEvents so Word.Application
' events are available for logging and interception.
' ============================================================

Private WithEvents m_oWord As Word.Application
Attribute m_oWord.VB_VarHelpID = -1
Private m_oLog As TextBox

Public Sub Export(oLog As TextBox, sOutputPath As String, sFilter As String, bFilterOn As Boolean)
    Set m_oLog = oLog
    APP.AppendLog m_oLog, "=== CNS Export - Word (VBA) started ==="

    On Error GoTo ErrHandler

    If Len(Trim(sOutputPath)) = 0 Then
        sOutputPath = APP.Get_ExportPath() & "ETW_cns_MsWord_VBA.docx"
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

    ' --- Open Word ---
    APP.AppendLog m_oLog, "  Opening Word..."
    Set m_oWord = New Word.Application
    m_oWord.Visible = True

    Dim oDoc As Word.Document
    Set oDoc = m_oWord.Documents.Add

    ' Title
    oDoc.Content.Text = "CNS Countries" & vbCrLf
    oDoc.Paragraphs(1).Style = oDoc.Styles(wdStyleHeading1)

    ' Table: header row + data rows (flag image column first)
    Dim oTable As Word.Table
    Set oTable = oDoc.Tables.Add(oDoc.Content, 1, 4)
    oTable.Style = "Table Grid"
    oTable.Cell(1, 1).Range.Text = "Flag"
    oTable.Cell(1, 2).Range.Text = "Country"
    oTable.Cell(1, 3).Range.Text = "ISO2"
    oTable.Cell(1, 4).Range.Text = "ISO3"

    Dim lRow As Long
    Dim sFlagPath As String
    Dim oPicture As Word.InlineShape
    lRow = 2
    Do While Not rs.EOF
        oTable.Rows.Add
        sFlagPath = ETW__FlagImages.Get_FlagImageFilePath( _
            m_oLog, Nz(rs.Fields("Iso2").Value, ""), Nz(rs.Fields("FlagUrl").Value, ""))
        If Len(sFlagPath) > 0 Then
            Set oPicture = oDoc.InlineShapes.AddPicture( _
                FileName:=sFlagPath, LinkToFile:=False, SaveWithDocument:=True, _
                Range:=oTable.Cell(lRow, 1).Range)
            oPicture.LockAspectRatio = msoTrue
            oPicture.Height = 14
        End If
        oTable.Cell(lRow, 2).Range.Text = Nz(rs.Fields("Country").Value, "")
        oTable.Cell(lRow, 3).Range.Text = Nz(rs.Fields("Iso2").Value, "")
        oTable.Cell(lRow, 4).Range.Text = Nz(rs.Fields("Iso3").Value, "")
        lRow = lRow + 1
        rs.MoveNext
    Loop

    rs.Close
    Set db = Nothing

    APP.AppendLog m_oLog, "  Saving to: " & sOutputPath
    oDoc.SaveAs2 sOutputPath, wdFormatXMLDocument
    oDoc.Close False
    m_oWord.Quit False
    Set oDoc = Nothing
    Set m_oWord = Nothing

    APP.AppendLog m_oLog, "=== CNS Export - Word (VBA) complete ==="
    Set m_oLog = Nothing
    Exit Sub

ErrHandler:
    APP.AppendLog m_oLog, "ERROR: " & Err.Description
    On Error Resume Next
    If Not oDoc Is Nothing Then oDoc.Close False
    If Not m_oWord Is Nothing Then m_oWord.Quit False
    Set m_oWord = Nothing
    Set m_oLog = Nothing
    On Error GoTo 0
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
' Word.Application events
' -------------------------------------------------------
Private Sub m_oWord_NewDocument(ByVal Doc As Word.Document)
    APP.AppendLog m_oLog, "  [Event] NewDocument: " & Doc.Name
End Sub

Private Sub m_oWord_DocumentBeforeSave(ByVal Doc As Word.Document, SaveAsUI As Boolean, Cancel As Boolean)
    APP.AppendLog m_oLog, "  [Event] DocumentBeforeSave: " & Doc.Name
End Sub

Private Sub m_oWord_DocumentBeforeClose(ByVal Doc As Word.Document, Cancel As Boolean)
    APP.AppendLog m_oLog, "  [Event] DocumentBeforeClose: " & Doc.Name
End Sub
