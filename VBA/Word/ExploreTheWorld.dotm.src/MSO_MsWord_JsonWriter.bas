Attribute VB_Name = "MSO_MsWord_JsonWriter"
Option Explicit

' Writes a Word document to a strictly valid JSON file.
'
' The output schema matches the JBC.ExploreTheWorld.DL.MsWord entity classes and is
' byte-identical to MsOfficeJsonSerializer output for the same object graph. The property
' names and values follow the VBA object model documented at
' https://learn.microsoft.com/en-us/office/vba/api/overview/
'
' JSON emission (comma handling, escaping, UTF-8 output) lives in MSO_JsonWriterCore.
' The blob options are accepted for signature parity with the other writers; the VBA
' object model offers no supported way to extract InlineShape image bytes, so the
' InlineShape Image property is always null here (the OpenXML writer provides the bytes).

Private m_oDocument As Word.Document

Public Sub WriteDocumentToJsonFile(oDocument As Word.Document, sOutputFilePath As String, _
    Optional eBlobOutput As JsonBlobOutput = jsonBlobBase64, Optional sBlobFolderPath As String = "")

    Dim dtStart As Date
    Dim dtStop As Date
    Dim iSeconds As Integer

    If (sOutputFilePath = "") Then
        Exit Sub
    End If

    dtStart = Now()

    Call LogStatus("Writing JSON: " & sOutputFilePath)

    Set m_oDocument = oDocument

    Call JsonWriter_Begin

    Call WriteDocument

    Call JsonWriter_End(sOutputFilePath)

    dtStop = Now()

    iSeconds = DateDiff("s", dtStart, dtStop)

    Call LogStatus("Done - " & iSeconds & " seconds.")
End Sub

' -- Document -----------------------------------------------------------------

Private Sub WriteDocument()
    Dim iLevel As Integer
    iLevel = 0

    ' Lists
    Call WriteBuiltInDocumentProperties(iLevel, m_oDocument.BuiltInDocumentProperties)
    Call WriteCustomDocumentProperties(iLevel, m_oDocument.CustomDocumentProperties)
    Call WriteBookmarks(iLevel, m_oDocument.Bookmarks)
    Call WriteComments(iLevel, m_oDocument.Comments)
    Call WriteContentControls(iLevel, m_oDocument.ContentControls)
    Call WriteEndnotes(iLevel, m_oDocument.Endnotes)
    Call WriteFields(iLevel, m_oDocument.Fields)
    Call WriteFootnotes(iLevel, m_oDocument.Footnotes)
    Call WriteInlineShapes(iLevel, m_oDocument.InlineShapes)
    Call WriteShapes(iLevel, m_oDocument.Shapes)
    Call WriteParagraphs(iLevel, m_oDocument.Paragraphs)
    Call WriteSections(iLevel, m_oDocument.Sections)
    Call WriteStyles(iLevel, m_oDocument.Styles)
    Call WriteTables(iLevel, m_oDocument.Tables)
    Call WriteVariables(iLevel, m_oDocument.Variables)

    ' Fields
    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", m_oDocument.Name)
    Call WritePropValueString(iLevel, "FullName", m_oDocument.FullName)
    Call WritePropValueString(iLevel, "Path", m_oDocument.Path)

    On Error Resume Next

    Call WritePropValueBoolean(iLevel, "Saved", m_oDocument.Saved)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Saved")
    End If

    Call WritePropValueLong(iLevel, "SaveFormat", m_oDocument.SaveFormat)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "SaveFormat")
    End If

    Call WritePropValueBoolean(iLevel, "ReadOnly", m_oDocument.ReadOnly)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ReadOnly")
    End If

    Call WritePropValueLong(iLevel, "ProtectionType", m_oDocument.ProtectionType)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ProtectionType")
    End If

    Call WritePropValueBoolean(iLevel, "TrackRevisions", m_oDocument.TrackRevisions)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TrackRevisions")
    End If

    Call WritePropValueBoolean(iLevel, "VBASigned", m_oDocument.VBASigned)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "VBASigned")
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

' -- Bookmarks ----------------------------------------------------------------

Private Sub WriteBookmarks(ByVal iLevel As Integer, oBookmarks As Word.Bookmarks)
    Dim oBookmark As Word.Bookmark

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "Bookmarks")

    For Each oBookmark In oBookmarks
        Call WriteBookmark(iLevel, oBookmark)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteBookmark(ByVal iLevel As Integer, oBookmark As Word.Bookmark)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", oBookmark.Name)

    On Error Resume Next
    Call WritePropValueString(iLevel, "RangeText", oBookmark.Range.Text)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "RangeText")
    End If
    On Error GoTo 0

    Call WritePropValueBoolean(iLevel, "Empty", oBookmark.Empty)

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Comments -----------------------------------------------------------------

Private Sub WriteComments(ByVal iLevel As Integer, oComments As Word.Comments)
    Dim oComment As Word.Comment

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "Comments")

    For Each oComment In oComments
        Call WriteComment(iLevel, oComment)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteComment(ByVal iLevel As Integer, oComment As Word.Comment)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "Index", oComment.Index)
    Call WritePropValueString(iLevel, "Author", oComment.Author)
    Call WritePropValueString(iLevel, "Initial", oComment.Initial)
    Call WritePropValueDate(iLevel, "Date", oComment.Date)

    On Error Resume Next
    Call WritePropValueString(iLevel, "RangeText", oComment.Range.Text)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "RangeText")
    End If
    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- ContentControls ----------------------------------------------------------

Private Sub WriteContentControls(ByVal iLevel As Integer, oContentControls As Word.ContentControls)
    Dim oCC As Word.ContentControl

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "ContentControls")

    For Each oCC In oContentControls
        Call WriteContentControl(iLevel, oCC)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteContentControl(ByVal iLevel As Integer, oCC As Word.ContentControl)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Title", oCC.Title)
    Call WritePropValueString(iLevel, "Tag", oCC.Tag)
    Call WritePropValueInteger(iLevel, "Type", oCC.Type)

    On Error Resume Next
    Call WritePropValueString(iLevel, "RangeText", oCC.Range.Text)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "RangeText")
    End If
    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Endnotes -----------------------------------------------------------------

Private Sub WriteEndnotes(ByVal iLevel As Integer, oEndnotes As Word.Endnotes)
    Dim oEndnote As Word.Endnote

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "Endnotes")

    For Each oEndnote In oEndnotes
        Call WriteEndnote(iLevel, oEndnote)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteEndnote(ByVal iLevel As Integer, oEndnote As Word.Endnote)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "Index", oEndnote.Index)

    On Error Resume Next
    Call WritePropValueString(iLevel, "RangeText", oEndnote.Range.Text)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "RangeText")
    End If
    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Fields -------------------------------------------------------------------

Private Sub WriteFields(ByVal iLevel As Integer, oFields As Word.Fields)
    Dim oField As Word.Field

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "Fields")

    For Each oField In oFields
        Call WriteField(iLevel, oField)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteField(ByVal iLevel As Integer, oField As Word.Field)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "Index", oField.Index)
    Call WritePropValueInteger(iLevel, "Type", oField.Type)
    Call WritePropValueBoolean(iLevel, "Locked", oField.Locked)

    On Error Resume Next
    Call WritePropValueString(iLevel, "Code", oField.Code.Text)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Code")
    End If

    Call WritePropValueString(iLevel, "Result", oField.Result.Text)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Result")
    End If
    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Footnotes ----------------------------------------------------------------

Private Sub WriteFootnotes(ByVal iLevel As Integer, oFootnotes As Word.Footnotes)
    Dim oFootnote As Word.Footnote

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "Footnotes")

    For Each oFootnote In oFootnotes
        Call WriteFootnote(iLevel, oFootnote)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteFootnote(ByVal iLevel As Integer, oFootnote As Word.Footnote)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "Index", oFootnote.Index)

    On Error Resume Next
    Call WritePropValueString(iLevel, "RangeText", oFootnote.Range.Text)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "RangeText")
    End If
    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- InlineShapes -------------------------------------------------------------

Private Sub WriteInlineShapes(ByVal iLevel As Integer, oInlineShapes As Word.InlineShapes)
    Dim oShape As Word.InlineShape

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "InlineShapes")

    For Each oShape In oInlineShapes
        Call WriteInlineShape(iLevel, oShape)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteInlineShape(ByVal iLevel As Integer, oShape As Word.InlineShape)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    ' Entities (image extraction is not supported by the VBA object model)
    Call WritePropValueNull(iLevel + 1, "Image")

    ' Fields
    iLevel = iLevel + 1

    Call WritePropValueInteger(iLevel, "Type", oShape.Type)
    Call WritePropValueSingleString(iLevel, "Width", oShape.Width)
    Call WritePropValueSingleString(iLevel, "Height", oShape.Height)

    On Error Resume Next
    Call WritePropValueString(iLevel, "AlternativeText", oShape.AlternativeText)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AlternativeText")
    End If
    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Shapes -------------------------------------------------------------------

Private Sub WriteShapes(ByVal iLevel As Integer, oShapes As Word.Shapes)
    Dim oShape As Word.Shape

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "Shapes")

    For Each oShape In oShapes
        Call WriteShape(iLevel, oShape)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteShape(ByVal iLevel As Integer, oShape As Word.Shape)
    Dim oShapeObj As Object
    Dim bHasTextFrame As Boolean

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", oShape.Name)
    Call WritePropValueInteger(iLevel, "Type", oShape.Type)
    Call WritePropValueSingleString(iLevel, "Left", oShape.Left)
    Call WritePropValueSingleString(iLevel, "Top", oShape.Top)
    Call WritePropValueSingleString(iLevel, "Width", oShape.Width)
    Call WritePropValueSingleString(iLevel, "Height", oShape.Height)
    Call WritePropValueSingleString(iLevel, "Rotation", oShape.Rotation)

    On Error Resume Next
    Call WritePropValueString(iLevel, "AlternativeText", oShape.AlternativeText)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AlternativeText")
    End If

    Set oShapeObj = oShape
    Err.Clear
    bHasTextFrame = oShapeObj.HasTextFrame
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HasTextFrame")
        Call WritePropValueNull(iLevel, "TextFrameText")
    ElseIf (bHasTextFrame = True) Then
        Call WritePropValueBoolean(iLevel, "HasTextFrame", True)
        Call WritePropValueString(iLevel, "TextFrameText", oShape.TextFrame.TextRange.Text)
        If (Err.Number <> 0) Then
            Err.Clear
            Call WritePropValueNull(iLevel, "TextFrameText")
        End If
    Else
        Call WritePropValueBoolean(iLevel, "HasTextFrame", False)
        Call WritePropValueNull(iLevel, "TextFrameText")
    End If
    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Paragraphs ---------------------------------------------------------------

Private Sub WriteParagraphs(ByVal iLevel As Integer, oParagraphs As Word.Paragraphs)
    Dim oParagraph As Word.Paragraph

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "Paragraphs")

    For Each oParagraph In oParagraphs
        Call WriteParagraph(iLevel, oParagraph)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteParagraph(ByVal iLevel As Integer, oParagraph As Word.Paragraph)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    ' Lists
    Call WriteWordRuns(iLevel, oParagraph.Range)

    ' Fields
    iLevel = iLevel + 1

    On Error Resume Next

    Call WritePropValueString(iLevel, "Text", oParagraph.Range.Text)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Text")
    End If

    Call WritePropValueString(iLevel, "Style", CStr(oParagraph.Style))
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Style")
    End If

    Call WritePropValueLong(iLevel, "Alignment", oParagraph.Alignment)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Alignment")
    End If

    Call WritePropValueSingleString(iLevel, "LeftIndent", oParagraph.LeftIndent)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "LeftIndent")
    End If

    Call WritePropValueSingleString(iLevel, "RightIndent", oParagraph.RightIndent)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "RightIndent")
    End If

    Call WritePropValueSingleString(iLevel, "FirstLineIndent", oParagraph.FirstLineIndent)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "FirstLineIndent")
    End If

    Call WritePropValueSingleString(iLevel, "SpaceBefore", oParagraph.SpaceBefore)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "SpaceBefore")
    End If

    Call WritePropValueSingleString(iLevel, "SpaceAfter", oParagraph.SpaceAfter)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "SpaceAfter")
    End If

    Call WritePropValueSingleString(iLevel, "LineSpacing", oParagraph.LineSpacing)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "LineSpacing")
    End If

    Call WritePropValueBoolean(iLevel, "KeepWithNext", oParagraph.KeepWithNext)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "KeepWithNext")
    End If

    Call WritePropValueBoolean(iLevel, "PageBreakBefore", oParagraph.PageBreakBefore)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "PageBreakBefore")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteWordRuns(ByVal iLevel As Integer, ByVal oRange As Word.Range)
    Dim oRun As Word.Range
    Dim iRunIndex As Long

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "Runs")

    For Each oRun In oRange.Characters
        iRunIndex = iRunIndex + 1
        Call WriteWordRun(iLevel, oRun, iRunIndex)
        ' Limit run count per paragraph to avoid excessively large output
        If (iRunIndex >= 500) Then
            Exit For
        End If
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteWordRun(ByVal iLevel As Integer, ByVal oChar As Word.Range, ByVal iIndex As Long)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    ' Entities
    Call WriteWordFont(iLevel, oChar.Font)

    ' Fields
    iLevel = iLevel + 1

    On Error Resume Next

    Call WritePropValueLong(iLevel, "Index", iIndex)

    Call WritePropValueString(iLevel, "Text", oChar.Text)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Text")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteWordFont(ByVal iLevel As Integer, ByVal oFont As Word.Font)
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

    Call WritePropValueLong(iLevel, "Bold", oFont.Bold)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Bold")
    End If

    Call WritePropValueLong(iLevel, "Italic", oFont.Italic)
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

    Call WritePropValueLong(iLevel, "StrikeThrough", oFont.StrikeThrough)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "StrikeThrough")
    End If

    Call WritePropValueLong(iLevel, "Superscript", oFont.Superscript)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Superscript")
    End If

    Call WritePropValueLong(iLevel, "Subscript", oFont.Subscript)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Subscript")
    End If

    Call WritePropValueSingleString(iLevel, "Scaling", oFont.Scaling)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Scaling")
    End If

    Call WritePropValueSingleString(iLevel, "Spacing", oFont.Spacing)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Spacing")
    End If

    Call WritePropValueSingleString(iLevel, "Position", oFont.Position)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Position")
    End If

    Call WritePropValueLong(iLevel, "Emboss", oFont.Emboss)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Emboss")
    End If

    Call WritePropValueLong(iLevel, "Shadow", oFont.Shadow)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Shadow")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Sections -----------------------------------------------------------------

Private Sub WriteSections(ByVal iLevel As Integer, oSections As Word.Sections)
    Dim oSection As Word.Section

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "Sections")

    For Each oSection In oSections
        Call WriteSection(iLevel, oSection)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteSection(ByVal iLevel As Integer, oSection As Word.Section)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    ' Entities
    Call WritePageSetup(iLevel, oSection.PageSetup)

    ' Fields
    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "Index", oSection.Index)

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WritePageSetup(ByVal iLevel As Integer, oPageSetup As Word.PageSetup)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "PageSetup")

    iLevel = iLevel + 1

    Call WritePropValueSingleString(iLevel, "PageWidth", oPageSetup.PageWidth)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "PageWidth")
    End If

    Call WritePropValueSingleString(iLevel, "PageHeight", oPageSetup.PageHeight)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "PageHeight")
    End If

    Call WritePropValueLong(iLevel, "Orientation", oPageSetup.Orientation)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Orientation")
    End If

    Call WritePropValueSingleString(iLevel, "TopMargin", oPageSetup.TopMargin)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TopMargin")
    End If

    Call WritePropValueSingleString(iLevel, "BottomMargin", oPageSetup.BottomMargin)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "BottomMargin")
    End If

    Call WritePropValueSingleString(iLevel, "LeftMargin", oPageSetup.LeftMargin)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "LeftMargin")
    End If

    Call WritePropValueSingleString(iLevel, "RightMargin", oPageSetup.RightMargin)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "RightMargin")
    End If

    Call WritePropValueSingleString(iLevel, "HeaderDistance", oPageSetup.HeaderDistance)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HeaderDistance")
    End If

    Call WritePropValueSingleString(iLevel, "FooterDistance", oPageSetup.FooterDistance)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "FooterDistance")
    End If

    Call WritePropValueLong(iLevel, "PaperSize", oPageSetup.PaperSize)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "PaperSize")
    End If

    Call WritePropValueLong(iLevel, "SectionStart", oPageSetup.SectionStart)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "SectionStart")
    End If

    Call WritePropValueBoolean(iLevel, "DifferentFirstPageHeaderFooter", oPageSetup.DifferentFirstPageHeaderFooter)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "DifferentFirstPageHeaderFooter")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Styles -------------------------------------------------------------------

Private Sub WriteStyles(ByVal iLevel As Integer, oStyles As Word.Styles)
    Dim oStyle As Word.Style

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "Styles")

    For Each oStyle In oStyles
        Call WriteStyle(iLevel, oStyle)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteStyle(ByVal iLevel As Integer, oStyle As Word.Style)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "NameLocal", oStyle.NameLocal)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NameLocal")
    End If

    Call WritePropValueLong(iLevel, "Type", oStyle.Type)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Type")
    End If

    Call WritePropValueBoolean(iLevel, "InUse", oStyle.InUse)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "InUse")
    End If

    Call WritePropValueBoolean(iLevel, "BuiltIn", oStyle.BuiltIn)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "BuiltIn")
    End If

    Call WritePropValueBoolean(iLevel, "AutomaticallyUpdate", oStyle.AutomaticallyUpdate)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AutomaticallyUpdate")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Tables -------------------------------------------------------------------

Private Sub WriteTables(ByVal iLevel As Integer, oTables As Word.Tables)
    Dim oTable As Word.Table

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "Tables")

    For Each oTable In oTables
        Call WriteTable(iLevel, oTable)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteTable(ByVal iLevel As Integer, oTable As Word.Table)
    Dim oRow As Word.Row

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    ' Lists
    Call WriteBeginObjectList(iLevel + 1, "Rows")

    For Each oRow In oTable.Rows
        Call WriteTableRow(iLevel + 1, oRow)
    Next

    Call WriteEndObjectList(iLevel + 1)

    ' Fields
    iLevel = iLevel + 1

    On Error Resume Next

    Call WritePropValueLong(iLevel, "RowCount", oTable.Rows.Count)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "RowCount")
    End If

    Call WritePropValueLong(iLevel, "ColumnCount", oTable.Columns.Count)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ColumnCount")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteTableRow(ByVal iLevel As Integer, oRow As Word.Row)
    Dim oCell As Word.Cell

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    ' Lists
    Call WriteBeginObjectList(iLevel + 1, "Cells")

    For Each oCell In oRow.Cells
        Call WriteTableCell(iLevel + 1, oCell)
    Next

    Call WriteEndObjectList(iLevel + 1)

    ' Fields
    Call WritePropValueLong(iLevel + 1, "Index", oRow.Index)

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteTableCell(ByVal iLevel As Integer, oCell As Word.Cell)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "RowIndex", oCell.RowIndex)
    Call WritePropValueLong(iLevel, "ColumnIndex", oCell.ColumnIndex)

    On Error Resume Next
    Call WritePropValueString(iLevel, "Text", oCell.Range.Text)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Text")
    End If
    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Variables ----------------------------------------------------------------

Private Sub WriteVariables(ByVal iLevel As Integer, oVariables As Word.Variables)
    Dim oVariable As Word.Variable

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "Variables")

    For Each oVariable In oVariables
        Call WriteVariable(iLevel, oVariable)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteVariable(ByVal iLevel As Integer, oVariable As Word.Variable)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", oVariable.Name)
    Call WritePropValueString(iLevel, "Value", CStr(oVariable.Value))

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Utilities ----------------------------------------------------------------

Private Sub LogStatus(sStatus As String)
    Debug.Print sStatus
End Sub
