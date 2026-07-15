VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ETW__cns_MsExcel_VBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' ETW__cns_MsExcel_VBA
' Exports cns_Country data to a Microsoft Excel workbook
' using the Excel COM object model (Office automation / VBA).
' Class module: m_oExcel is WithEvents so Excel.Application
' events are available for logging and interception.
' ============================================================

Private WithEvents m_oExcel As Excel.Application
Attribute m_oExcel.VB_VarHelpID = -1
Private m_oLog As TextBox

Public Sub Export(oLog As TextBox, sOutputPath As String, sFilter As String, bFilterOn As Boolean)
    Set m_oLog = oLog
    APP.AppendLog m_oLog, "=== CNS Export - Excel (VBA) started ==="

    On Error GoTo ErrHandler

    If Len(Trim(sOutputPath)) = 0 Then
        sOutputPath = APP.Get_ExportPath() & "ETW_cns_MsExcel_VBA.xlsx"
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

    ' --- Open Excel ---
    APP.AppendLog m_oLog, "  Opening Excel..."
    Set m_oExcel = New Excel.Application
    m_oExcel.Visible = True

    Dim oWb As Excel.Workbook
    Set oWb = m_oExcel.Workbooks.Add

    Dim oWs As Excel.Worksheet
    Set oWs = oWb.Worksheets(1)
    oWs.Name = "CNS Countries"

    ' Header row (flag image column first)
    oWs.Cells(1, 1).Value = "Flag"
    oWs.Cells(1, 2).Value = "Country"
    oWs.Cells(1, 3).Value = "ISO2"
    oWs.Cells(1, 4).Value = "ISO3"
    oWs.Rows(1).Font.Bold = True
    oWs.Columns(1).ColumnWidth = 7

    Dim lRow As Long
    Dim sFlagPath As String
    Dim oCell As Excel.Range
    Dim oPicture As Excel.Shape
    lRow = 2
    Do While Not rs.EOF
        sFlagPath = ETW__FlagImages.Get_FlagImageFilePath( _
            m_oLog, Nz(rs.Fields("Iso2").Value, ""), Nz(rs.Fields("FlagUrl").Value, ""))
        If Len(sFlagPath) > 0 Then
            Set oCell = oWs.Cells(lRow, 1)
            oCell.RowHeight = 20
            ' Insert at original size, then scale by height with the aspect ratio locked.
            Set oPicture = oWs.Shapes.AddPicture(sFlagPath, msoFalse, msoCTrue, _
                oCell.Left + 2, oCell.Top + 2, -1, -1)
            oPicture.LockAspectRatio = msoTrue
            oPicture.Height = 16
        End If
        oWs.Cells(lRow, 2).Value = Nz(rs.Fields("Country").Value, "")
        oWs.Cells(lRow, 3).Value = Nz(rs.Fields("Iso2").Value, "")
        oWs.Cells(lRow, 4).Value = Nz(rs.Fields("Iso3").Value, "")
        lRow = lRow + 1
        rs.MoveNext
    Loop

    oWs.Columns("B:D").AutoFit

    rs.Close
    Set db = Nothing

    APP.AppendLog m_oLog, "  Saving to: " & sOutputPath
    oWb.SaveAs sOutputPath, xlOpenXMLWorkbook
    oWb.Close False
    m_oExcel.Quit
    Set oWs = Nothing
    Set oWb = Nothing
    Set m_oExcel = Nothing

    APP.AppendLog m_oLog, "=== CNS Export - Excel (VBA) complete ==="
    Set m_oLog = Nothing
    Exit Sub

ErrHandler:
    APP.AppendLog m_oLog, "ERROR: " & Err.Description
    On Error Resume Next
    If Not oWb Is Nothing Then oWb.Close False
    If Not m_oExcel Is Nothing Then m_oExcel.Quit
    Set m_oExcel = Nothing
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
' Excel.Application events
' -------------------------------------------------------
Private Sub m_oExcel_NewWorkbook(ByVal Wb As Excel.Workbook)
    APP.AppendLog m_oLog, "  [Event] NewWorkbook: " & Wb.Name
End Sub

Private Sub m_oExcel_WorkbookBeforeClose(ByVal Wb As Excel.Workbook, Cancel As Boolean)
    APP.AppendLog m_oLog, "  [Event] WorkbookBeforeClose: " & Wb.Name
End Sub

Private Sub m_oExcel_WorkbookBeforeSave(ByVal Wb As Excel.Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    APP.AppendLog m_oLog, "  [Event] WorkbookBeforeSave: " & Wb.Name
End Sub
