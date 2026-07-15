Attribute VB_Name = "MSO_MsExcel_JsonWriter"
Option Compare Database

Option Explicit

' Writes an Excel workbook to a strictly valid JSON file.
'
' The output schema matches the JBC.ExploreTheWorld.DL.MsExcel entity classes and is
' byte-identical to MsOfficeJsonSerializer output for the same object graph. The property
' names and values follow the VBA object model documented at
' https://learn.microsoft.com/en-us/office/vba/api/overview/
'
' JSON emission (comma handling, escaping, UTF-8 output) lives in MSO_JsonWriterCore.
' The blob options are accepted for signature parity with the other writers; the Excel
' schema currently has no image/blob properties.

Private m_oWorkbook As Excel.Workbook

Public Sub WriteWorkbookToJsonFile(oWorkbook As Excel.Workbook, sOutputFilePath As String, _
    Optional eBlobOutput As JsonBlobOutput = jsonBlobBase64, Optional sBlobFolderPath As String = "")

    Dim dtStart As Date
    Dim dtStop As Date
    Dim iSeconds As Integer

    If (sOutputFilePath = "") Then
        Exit Sub
    End If

    dtStart = Now()

    Call LogStatus("Writing JSON: " & sOutputFilePath)

    Set m_oWorkbook = oWorkbook

    Call JsonWriter_Begin

    Call WriteWorkbook

    Call JsonWriter_End(sOutputFilePath)

    dtStop = Now()

    iSeconds = DateDiff("s", dtStart, dtStop)

    Call LogStatus("Done - " & iSeconds & " seconds.")
End Sub

' -- Workbook -----------------------------------------------------------------

Private Sub WriteWorkbook()
    Dim iLevel As Integer
    iLevel = 0

    ' Lists
    Call WriteBuiltInDocumentProperties(iLevel, m_oWorkbook.BuiltInDocumentProperties)
    Call WriteCustomDocumentProperties(iLevel, m_oWorkbook.CustomDocumentProperties)
    Call WriteSheets(iLevel, m_oWorkbook.Sheets)

    ' Fields
    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", m_oWorkbook.Name)
    Call WritePropValueString(iLevel, "FullName", m_oWorkbook.FullName)
    Call WritePropValueString(iLevel, "Path", m_oWorkbook.Path)

    On Error Resume Next

    Call WritePropValueBoolean(iLevel, "Saved", m_oWorkbook.Saved)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Saved")
    End If

    Call WritePropValueBoolean(iLevel, "ReadOnly", m_oWorkbook.ReadOnly)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ReadOnly")
    End If

    Call WritePropValueBoolean(iLevel, "HasVBProject", m_oWorkbook.HasVBProject)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HasVBProject")
    End If

    Call WritePropValueLong(iLevel, "FileFormat", m_oWorkbook.FileFormat)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "FileFormat")
    End If

    Call WritePropValueString(iLevel, "CodeName", m_oWorkbook.CodeName)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "CodeName")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1
End Sub

' -- Document properties --------------------------------------------------------

Private Sub WriteBuiltInDocumentProperties(ByVal iLevel As Integer, oBuiltInDocumentProperties As Object)
    Dim p As Object

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "BuiltInDocumentProperties")

    For Each p In oBuiltInDocumentProperties
        Call WriteDocumentProperty(iLevel, p)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteCustomDocumentProperties(ByVal iLevel As Integer, oCustomDocumentProperties As Object)
    Dim p As Object

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "CustomDocumentProperties")

    For Each p In oCustomDocumentProperties
        Call WriteDocumentProperty(iLevel, p)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteDocumentProperty(ByVal iLevel As Integer, oProperty As Object)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", oProperty.Name)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Name")
    End If

    Call WritePropValueLong(iLevel, "Creator", oProperty.Creator)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Creator")
    End If

    Call WritePropValueString(iLevel, "LinkSource", oProperty.LinkSource)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "LinkSource")
    End If

    Call WritePropValueInteger(iLevel, "LinkToContent", oProperty.LinkToContent)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "LinkToContent")
    End If

    Call WritePropValueInteger(iLevel, "Type", oProperty.Type)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Type")
    End If

    Call WritePropValueString(iLevel, "Value", oProperty.Value)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Value")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Sheets -------------------------------------------------------------------

Private Sub WriteSheets(ByVal iLevel As Integer, oSheets As Object)
    Dim oSheet As Object

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "Sheets")

    For Each oSheet In oSheets
        Call WriteSheet(iLevel, oSheet)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteSheet(ByVal iLevel As Integer, oSheet As Object)
    Dim bIsWorksheet As Boolean

    bIsWorksheet = (TypeName(oSheet) = "Worksheet")

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    ' Entities and lists
    If (bIsWorksheet = True) Then
        Call WriteUsedRangeAndRows(iLevel, oSheet)
    Else
        Call WritePropValueNull(iLevel + 1, "UsedRange")
        Call WritePropValueNull(iLevel + 1, "Rows")
    End If

    ' Fields
    iLevel = iLevel + 1

    On Error Resume Next

    Call WritePropValueString(iLevel, "Name", oSheet.Name)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Name")
    End If

    Call WritePropValueLong(iLevel, "Index", oSheet.Index)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Index")
    End If

    Call WritePropValueString(iLevel, "CodeName", oSheet.CodeName)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "CodeName")
    End If

    Call WritePropValueLong(iLevel, "Visible", oSheet.Visible)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Visible")
    End If

    If (bIsWorksheet = True) Then
        Call WritePropValueString(iLevel, "StandardWidth", CStr(oSheet.StandardWidth))
        If (Err.Number <> 0) Then
            Err.Clear
            Call WritePropValueNull(iLevel, "StandardWidth")
        End If
    Else
        Call WritePropValueNull(iLevel, "StandardWidth")
    End If

    Call WritePropValueString(iLevel, "Type", TypeName(oSheet))

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteUsedRangeAndRows(ByVal iLevel As Integer, oSheet As Excel.Worksheet)
    Dim oUsedRange As Excel.Range
    Dim iRowCount As Long
    Dim iColCount As Long
    Dim iRow As Long

    On Error Resume Next
    Set oUsedRange = oSheet.UsedRange
    If (Err.Number <> 0) Then
        Err.Clear
        On Error GoTo 0
        Call WritePropValueNull(iLevel + 1, "UsedRange")
        Call WriteBeginEndObjectList(iLevel + 1, "Rows")
        Exit Sub
    End If
    On Error GoTo 0

    iRowCount = oUsedRange.Rows.Count
    iColCount = oUsedRange.Columns.Count

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "UsedRange")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Address", oUsedRange.Address)
    Call WritePropValueLong(iLevel, "RowCount", iRowCount)
    Call WritePropValueLong(iLevel, "ColumnCount", iColCount)

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)

    Call WriteBeginObjectList(iLevel, "Rows")

    For iRow = 1 To iRowCount
        Call WriteWorksheetRow(iLevel, oUsedRange.Rows(iRow), iRow, iColCount)
    Next iRow

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteWorksheetRow(ByVal iLevel As Integer, oRow As Excel.Range, iRowIndex As Long, iColCount As Long)
    Dim iCol As Long

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    Call WriteBeginObjectList(iLevel + 1, "Cells")

    For iCol = 1 To iColCount
        Call WriteCell(iLevel + 1, oRow.Cells(1, iCol), iRowIndex, iCol)
    Next iCol

    Call WriteEndObjectList(iLevel + 1)

    Call WritePropValueLong(iLevel + 1, "RowIndex", iRowIndex)

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteCell(ByVal iLevel As Integer, oCell As Excel.Range, iRowIndex As Long, iColIndex As Long)
    Dim oComment As Excel.Comment

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    ' Entities
    Call WriteCellFont(iLevel, oCell.Font)
    Call WriteCellInterior(iLevel, oCell.Interior)

    ' Fields
    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "RowIndex", iRowIndex)
    Call WritePropValueLong(iLevel, "ColumnIndex", iColIndex)
    Call WritePropValueString(iLevel, "Address", oCell.Address)

    On Error Resume Next

    Call WritePropValueString(iLevel, "Value", CStr(oCell.Value))
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Value")
    End If

    Call WritePropValueString(iLevel, "Formula", CStr(oCell.Formula))
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Formula")
    End If

    Call WritePropValueString(iLevel, "NumberFormat", CStr(oCell.NumberFormat))
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NumberFormat")
    End If

    Call WritePropValueString(iLevel, "Text", CStr(oCell.Text))
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Text")
    End If

    Call WritePropValueLong(iLevel, "HorizontalAlignment", oCell.HorizontalAlignment)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HorizontalAlignment")
    End If

    Call WritePropValueLong(iLevel, "VerticalAlignment", oCell.VerticalAlignment)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "VerticalAlignment")
    End If

    Call WritePropValueBoolean(iLevel, "WrapText", oCell.WrapText)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "WrapText")
    End If

    Call WritePropValueBoolean(iLevel, "MergeCells", oCell.MergeCells)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "MergeCells")
    End If

    ' Comment
    Err.Clear
    Set oComment = oCell.Comment
    If (Err.Number = 0) And Not (oComment Is Nothing) Then
        Call WritePropValueBoolean(iLevel, "HasComment", True)
        Call WritePropValueString(iLevel, "Comment", CStr(oComment.Text))
        If (Err.Number <> 0) Then
            Err.Clear
            Call WritePropValueNull(iLevel, "Comment")
        End If
    Else
        Err.Clear
        Call WritePropValueBoolean(iLevel, "HasComment", False)
        Call WritePropValueNull(iLevel, "Comment")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteCellFont(ByVal iLevel As Integer, oFont As Excel.Font)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "Font")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", oFont.Name)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Name")
    End If

    Call WritePropValueSingleString(iLevel, "Size", oFont.Size)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Size")
    End If

    Call WritePropValueBoolean(iLevel, "Bold", oFont.Bold)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Bold")
    End If

    Call WritePropValueBoolean(iLevel, "Italic", oFont.Italic)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Italic")
    End If

    Call WritePropValueLong(iLevel, "Underline", oFont.Underline)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Underline")
    End If

    Call WritePropValueLong(iLevel, "Color", oFont.Color)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Color")
    End If

    Call WritePropValueLong(iLevel, "ColorIndex", oFont.ColorIndex)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ColorIndex")
    End If

    Call WritePropValueBoolean(iLevel, "StrikeThrough", oFont.StrikeThrough)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "StrikeThrough")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteCellInterior(ByVal iLevel As Integer, oInterior As Excel.Interior)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "Interior")

    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "Color", oInterior.Color)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Color")
    End If

    Call WritePropValueLong(iLevel, "ColorIndex", oInterior.ColorIndex)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ColorIndex")
    End If

    Call WritePropValueLong(iLevel, "Pattern", oInterior.Pattern)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Pattern")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Utilities ----------------------------------------------------------------

Private Sub LogStatus(sStatus As String)
    Debug.Print sStatus
End Sub
