Attribute VB_Name = "MSO_MsPowerPoint_JsonWriter"
Option Compare Database

Option Explicit

' Writes a PowerPoint presentation to a strictly valid JSON file.
'
' The output schema matches the JBC.ExploreTheWorld.DL.MsPowerPoint entity classes and is
' byte-identical to MsOfficeJsonSerializer output for the same object graph. The property
' names and values follow the VBA object model documented at
' https://learn.microsoft.com/en-us/office/vba/api/overview/
'
' JSON emission (comma handling, escaping, UTF-8 output) lives in MSO_JsonWriterCore.

Private m_oPresentation As PowerPoint.Presentation
Private m_eBlobOutput As JsonBlobOutput
Private m_sBlobFolderPath As String
Private m_sBlobFolderName As String

Public Sub WritePresentationToJsonFile(oPresentation As PowerPoint.Presentation, sOutputFilePath As String, _
    Optional eBlobOutput As JsonBlobOutput = jsonBlobBase64, Optional sBlobFolderPath As String = "")

    Dim dtStart As Date
    Dim dtStop As Date
    Dim iSeconds As Integer
    Dim oFSO As Scripting.FileSystemObject

    If (sOutputFilePath = "") Then
        Exit Sub
    End If

    dtStart = Now()

    Call LogStatus("Writing JSON: " & sOutputFilePath)

    Set m_oPresentation = oPresentation

    m_eBlobOutput = eBlobOutput
    m_sBlobFolderPath = GetBlobFolderPath(sOutputFilePath, sBlobFolderPath)

    Set oFSO = New Scripting.FileSystemObject
    m_sBlobFolderName = oFSO.GetFileName(m_sBlobFolderPath)

    Call JsonWriter_Begin

    Call WritePresentation

    Call JsonWriter_End(sOutputFilePath)

    dtStop = Now()

    iSeconds = DateDiff("s", dtStart, dtStop)

    Call LogStatus("Done - " & iSeconds & " seconds.")
End Sub

Private Sub WritePresentation()
    Dim iLevel As Integer
    iLevel = 0

    ' Entities
    Call WriteMasterSafe(iLevel, "SlideMaster")
    Call WriteMasterSafe(iLevel, "TitleMaster")
    Call WriteMasterSafe(iLevel, "HandoutMaster")
    Call WriteMasterSafe(iLevel, "NotesMaster")

    Call WriteCoauthoring(iLevel, m_oPresentation.Coauthoring)

    Call WritePageSetup(iLevel, m_oPresentation.PageSetup)

    Call WritePrintOptions(iLevel, m_oPresentation.PrintOptions)

    Call WriteSlideShowSettings(iLevel, m_oPresentation.SlideShowSettings)

    ' Lists
    Call WriteBuiltInDocumentProperties(iLevel, m_oPresentation.BuiltInDocumentProperties)

    Call WriteCustomDocumentProperties(iLevel, m_oPresentation.CustomDocumentProperties)

    On Error Resume Next
    Call WriteContentTypeProperties(iLevel, m_oPresentation.ContentTypeProperties)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "ContentTypeProperties")
    End If
    On Error GoTo 0

    Call WriteColorSchemes(iLevel)

    Call WriteCustomXmlParts(iLevel)

    Call WriteDesigns(iLevel)

    Call WriteExtraColors(iLevel)

    Call WriteFonts(iLevel)

    Call WriteGuides(iLevel)

    Call WriteSectionProperties(iLevel, m_oPresentation.SectionProperties)

    Call WriteSlides(iLevel, m_oPresentation.Slides)

    Call WriteTags(iLevel, m_oPresentation.Tags)

    ' Fields
    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", m_oPresentation.Name)
    Call WritePropValueString(iLevel, "Path", m_oPresentation.Path)
    Call WritePropValueString(iLevel, "FullName", m_oPresentation.FullName)

    On Error Resume Next

    Call WritePropValueBoolean(iLevel, "AutoSaveOn", m_oPresentation.AutoSaveOn)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AutoSaveOn")
    End If

    Call WritePropValueBoolean(iLevel, "ChartDataPointTrack", m_oPresentation.ChartDataPointTrack)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ChartDataPointTrack")
    End If

    Call WritePropValueLong(iLevel, "CreateVideoStatus", m_oPresentation.CreateVideoStatus)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "CreateVideoStatus")
    End If

    Call WritePropValueLong(iLevel, "DefaultLanguageID", m_oPresentation.DefaultLanguageID)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "DefaultLanguageID")
    End If

    Call WritePropValueLong(iLevel, "DisplayComments", m_oPresentation.DisplayComments)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "DisplayComments")
    End If

    Call WritePropValueString(iLevel, "EncryptionProvider", m_oPresentation.EncryptionProvider)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "EncryptionProvider")
    End If

    Call WritePropValueLong(iLevel, "EnvelopeVisible", m_oPresentation.EnvelopeVisible)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "EnvelopeVisible")
    End If

    Call WritePropValueLong(iLevel, "FarEastLineBreakLanguage", m_oPresentation.FarEastLineBreakLanguage)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "FarEastLineBreakLanguage")
    End If

    Call WritePropValueLong(iLevel, "FarEastLineBreakLevel", m_oPresentation.FarEastLineBreakLevel)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "FarEastLineBreakLevel")
    End If

    Call WritePropValueBoolean(iLevel, "Final", m_oPresentation.Final)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Final")
    End If

    Call WritePropValueSingleString(iLevel, "GridDistance", m_oPresentation.GridDistance)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "GridDistance")
    End If

    Call WritePropValueLong(iLevel, "HasHandoutMaster", m_oPresentation.HasHandoutMaster)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HasHandoutMaster")
    End If

    Call WritePropValueLong(iLevel, "HasNotesMaster", m_oPresentation.HasNotesMaster)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HasNotesMaster")
    End If

    Call WritePropValueLong(iLevel, "HasTitleMaster", m_oPresentation.HasTitleMaster)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HasTitleMaster")
    End If

    Call WritePropValueLong(iLevel, "HasVBProject", m_oPresentation.HasVBProject)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HasVBProject")
    End If

    Call WritePropValueBoolean(iLevel, "InMergeMode", m_oPresentation.InMergeMode)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "InMergeMode")
    End If

    Call WritePropValueLong(iLevel, "LayoutDirection", m_oPresentation.LayoutDirection)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "LayoutDirection")
    End If

    Call WritePropValueString(iLevel, "NoLineBreakAfter", m_oPresentation.NoLineBreakAfter)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NoLineBreakAfter")
    End If

    Call WritePropValueString(iLevel, "NoLineBreakBefore", m_oPresentation.NoLineBreakBefore)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NoLineBreakBefore")
    End If

    Call WritePropValueString(iLevel, "Password", m_oPresentation.Password)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Password")
    End If

    Call WritePropValueString(iLevel, "PasswordEncryptionAlgorithm", m_oPresentation.PasswordEncryptionAlgorithm)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "PasswordEncryptionAlgorithm")
    End If

    Call WritePropValueBoolean(iLevel, "PasswordEncryptionFileProperties", m_oPresentation.PasswordEncryptionFileProperties)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "PasswordEncryptionFileProperties")
    End If

    Call WritePropValueLong(iLevel, "PasswordEncryptionKeyLength", m_oPresentation.PasswordEncryptionKeyLength)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "PasswordEncryptionKeyLength")
    End If

    Call WritePropValueString(iLevel, "PasswordEncryptionProvider", m_oPresentation.PasswordEncryptionProvider)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "PasswordEncryptionProvider")
    End If

    Call WritePropValueLong(iLevel, "ReadOnly", m_oPresentation.ReadOnly)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ReadOnly")
    End If

    Call WritePropValueBoolean(iLevel, "ReadOnlyRecommended", m_oPresentation.ReadOnlyRecommended)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ReadOnlyRecommended")
    End If

    Call WritePropValueLong(iLevel, "RemovePersonalInformation", m_oPresentation.RemovePersonalInformation)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "RemovePersonalInformation")
    End If

    Call WritePropValueLong(iLevel, "Saved", m_oPresentation.Saved)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Saved")
    End If

    Call WritePropValueLong(iLevel, "SnapToGrid", m_oPresentation.SnapToGrid)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "SnapToGrid")
    End If

    Call WritePropValueString(iLevel, "TemplateName", m_oPresentation.TemplateName)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TemplateName")
    End If

    Call WritePropValueLong(iLevel, "VBASigned", m_oPresentation.VBASigned)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "VBASigned")
    End If

    Call WritePropValueString(iLevel, "WritePassword", m_oPresentation.WritePassword)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "WritePassword")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1
End Sub

' -- Masters --------------------------------------------------------------------

Private Sub WriteMasterSafe(ByVal iLevel As Integer, sNameAs As String)
    Dim oMaster As PowerPoint.Master
    Dim bIsSlideMaster As Boolean

    bIsSlideMaster = (sNameAs = "SlideMaster")

    On Error Resume Next

    Select Case sNameAs
        Case "SlideMaster"
            Set oMaster = m_oPresentation.SlideMaster
        Case "TitleMaster"
            If (m_oPresentation.HasTitleMaster = msoTrue) Then
                Set oMaster = m_oPresentation.TitleMaster
            End If
        Case "HandoutMaster"
            Set oMaster = m_oPresentation.HandoutMaster
        Case "NotesMaster"
            Set oMaster = m_oPresentation.NotesMaster
    End Select

    If (Err.Number <> 0) Or (oMaster Is Nothing) Then
        Err.Clear
        On Error GoTo 0
        Call WritePropValueNull(iLevel + 1, sNameAs)
        Exit Sub
    End If

    On Error GoTo 0

    Call WriteMaster(iLevel, oMaster, sNameAs, bIsSlideMaster)
End Sub

Private Sub WriteMaster(ByVal iLevel As Integer, oMaster As PowerPoint.Master, sNameAs As String, bIsSlideMaster As Boolean)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, sNameAs)

    On Error Resume Next
    Call WriteSlideShowTransition(iLevel, oMaster.SlideShowTransition)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "SlideShowTransition")
    End If
    On Error GoTo 0

    If (bIsSlideMaster = True) Then
        Call WriteSlideMasterLayouts(iLevel, oMaster.CustomLayouts)
    Else
        Call WritePropValueNull(iLevel + 1, "SlideMasterLayouts")
    End If

    Call WriteShapes(iLevel, oMaster.Shapes)

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", oMaster.Name)

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteSlideMasterLayouts(ByVal iLevel As Integer, oCustomLayouts As PowerPoint.CustomLayouts)
    Dim oCustomLayout As PowerPoint.CustomLayout

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "SlideMasterLayouts")

    For Each oCustomLayout In oCustomLayouts
        Call WriteSlideMasterLayout(iLevel, oCustomLayout)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteSlideMasterLayout(ByVal iLevel As Integer, oCustomLayout As PowerPoint.CustomLayout)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    On Error Resume Next
    Call WriteSlideShowTransition(iLevel, oCustomLayout.SlideShowTransition)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "SlideShowTransition")
    End If
    On Error GoTo 0

    Call WriteShapes(iLevel, oCustomLayout.Shapes)

    Call WriteHyperlinks(iLevel, oCustomLayout.Hyperlinks, "Hyperlinks")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", oCustomLayout.Name)
    Call WritePropValueLong(iLevel, "Index", oCustomLayout.Index)

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Coauthoring ----------------------------------------------------------------

Private Sub WriteCoauthoring(ByVal iLevel As Integer, oCoauthoring As PowerPoint.Coauthoring)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "Coauthoring")

    iLevel = iLevel + 1

    Call WritePropValueBoolean(iLevel, "FavorServerEditsDuringMerge", oCoauthoring.FavorServerEditsDuringMerge)
    Call WritePropValueBoolean(iLevel, "MergeMode", oCoauthoring.MergeMode)
    Call WritePropValueBoolean(iLevel, "PendingUpdates", oCoauthoring.PendingUpdates)

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- PageSetup ------------------------------------------------------------------

Private Sub WritePageSetup(ByVal iLevel As Integer, oPageSetup As PowerPoint.PageSetup)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "PageSetup")

    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "FirstSlideNumber", oPageSetup.FirstSlideNumber)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "FirstSlideNumber")
    End If

    Call WritePropValueLong(iLevel, "NotesOrientation", oPageSetup.NotesOrientation)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NotesOrientation")
    End If

    Call WritePropValueSingleString(iLevel, "SlideHeight", oPageSetup.SlideHeight)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "SlideHeight")
    End If

    Call WritePropValueLong(iLevel, "SlideOrientation", oPageSetup.SlideOrientation)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "SlideOrientation")
    End If

    Call WritePropValueLong(iLevel, "SlideSize", oPageSetup.SlideSize)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "SlideSize")
    End If

    Call WritePropValueSingleString(iLevel, "SlideWidth", oPageSetup.SlideWidth)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "SlideWidth")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- PrintOptions ---------------------------------------------------------------

Private Sub WritePrintOptions(ByVal iLevel As Integer, oPrintOptions As PowerPoint.PrintOptions)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "PrintOptions")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "ActivePrinter", oPrintOptions.ActivePrinter)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ActivePrinter")
    End If

    Call WritePropValueLong(iLevel, "Collate", oPrintOptions.Collate)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Collate")
    End If

    Call WritePropValueLong(iLevel, "FitToPage", oPrintOptions.FitToPage)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "FitToPage")
    End If

    Call WritePropValueLong(iLevel, "FrameSlides", oPrintOptions.FrameSlides)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "FrameSlides")
    End If

    Call WritePropValueLong(iLevel, "HandoutOrder", oPrintOptions.HandoutOrder)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HandoutOrder")
    End If

    Call WritePropValueLong(iLevel, "HighQuality", oPrintOptions.HighQuality)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HighQuality")
    End If

    Call WritePropValueLong(iLevel, "NumberOfCopies", oPrintOptions.NumberOfCopies)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NumberOfCopies")
    End If

    Call WritePropValueLong(iLevel, "OutputType", oPrintOptions.OutputType)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "OutputType")
    End If

    Call WritePropValueLong(iLevel, "PrintColorType", oPrintOptions.PrintColorType)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "PrintColorType")
    End If

    Call WritePropValueLong(iLevel, "PrintComments", oPrintOptions.PrintComments)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "PrintComments")
    End If

    Call WritePropValueLong(iLevel, "PrintFontsAsGraphics", oPrintOptions.PrintFontsAsGraphics)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "PrintFontsAsGraphics")
    End If

    Call WritePropValueLong(iLevel, "PrintHiddenSlides", oPrintOptions.PrintHiddenSlides)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "PrintHiddenSlides")
    End If

    Call WritePropValueLong(iLevel, "PrintInBackground", oPrintOptions.PrintInBackground)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "PrintInBackground")
    End If

    Call WritePropValueLong(iLevel, "RangeType", oPrintOptions.RangeType)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "RangeType")
    End If

    Call WritePropValueString(iLevel, "SlideShowName", oPrintOptions.SlideShowName)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "SlideShowName")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- SlideShowSettings ----------------------------------------------------------

Private Sub WriteSlideShowSettings(ByVal iLevel As Integer, oSettings As PowerPoint.SlideShowSettings)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "SlideShowSettings")

    Call WriteColorFormat(iLevel, oSettings.PointerColor, "PointerColor")
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "PointerColor")
    End If

    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "AdvanceMode", oSettings.AdvanceMode)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AdvanceMode")
    End If

    Call WritePropValueLong(iLevel, "EndingSlide", oSettings.EndingSlide)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "EndingSlide")
    End If

    Call WritePropValueLong(iLevel, "LoopUntilStopped", oSettings.LoopUntilStopped)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "LoopUntilStopped")
    End If

    Call WritePropValueLong(iLevel, "RangeType", oSettings.RangeType)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "RangeType")
    End If

    Call WritePropValueLong(iLevel, "ShowMediaControls", oSettings.ShowMediaControls)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ShowMediaControls")
    End If

    Call WritePropValueLong(iLevel, "ShowPresenterView", oSettings.ShowPresenterView)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ShowPresenterView")
    End If

    Call WritePropValueLong(iLevel, "ShowScrollbar", oSettings.ShowScrollbar)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ShowScrollbar")
    End If

    Call WritePropValueLong(iLevel, "ShowType", oSettings.ShowType)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ShowType")
    End If

    Call WritePropValueLong(iLevel, "ShowWithAnimation", oSettings.ShowWithAnimation)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ShowWithAnimation")
    End If

    Call WritePropValueLong(iLevel, "ShowWithNarration", oSettings.ShowWithNarration)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ShowWithNarration")
    End If

    Call WritePropValueLong(iLevel, "StartingSlide", oSettings.StartingSlide)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "StartingSlide")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Document properties ----------------------------------------------------------

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
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", oProperty.Name)
    Call WritePropValueLong(iLevel, "Creator", oProperty.Creator)

    On Error Resume Next
    Call WritePropValueString(iLevel, "LinkSource", oProperty.LinkSource)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "LinkSource")
    End If
    On Error GoTo 0

    Call WritePropValueInteger(iLevel, "LinkToContent", oProperty.LinkToContent)
    Call WritePropValueInteger(iLevel, "Type", oProperty.Type)

    On Error Resume Next
    Call WritePropValueString(iLevel, "Value", oProperty.Value)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Value")
    End If
    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteContentTypeProperties(ByVal iLevel As Integer, oContentTypeProperties As MetaProperties)
    Dim p As Object

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "ContentTypeProperties")

    For Each p In oContentTypeProperties
        Call WriteContentTypeProperty(iLevel, p)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteContentTypeProperty(ByVal iLevel As Integer, oContentTypeProperty As MetaProperty)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Id", oContentTypeProperty.Id)
    Call WritePropValueString(iLevel, "Name", oContentTypeProperty.Name)
    Call WritePropValueLong(iLevel, "Creator", oContentTypeProperty.Creator)
    Call WritePropValueBoolean(iLevel, "IsReadOnly", oContentTypeProperty.IsReadOnly)
    Call WritePropValueBoolean(iLevel, "IsRequired", oContentTypeProperty.IsRequired)
    Call WritePropValueInteger(iLevel, "Type", oContentTypeProperty.Type)

    On Error Resume Next
    Call WritePropValueString(iLevel, "Value", oContentTypeProperty.Value)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Value")
    End If
    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Presentation collections -----------------------------------------------------

Private Sub WriteColorSchemes(ByVal iLevel As Integer)
    Dim oColorScheme As PowerPoint.ColorScheme

    iLevel = iLevel + 1

    On Error Resume Next

    If (m_oPresentation.ColorSchemes.Count < 0) Then
    End If

    If (Err.Number <> 0) Then
        Err.Clear
        On Error GoTo 0
        Call WritePropValueNull(iLevel, "ColorSchemes")
        Exit Sub
    End If

    On Error GoTo 0

    Call WriteBeginObjectList(iLevel, "ColorSchemes")

    For Each oColorScheme In m_oPresentation.ColorSchemes
        Call WriteColorScheme(iLevel, oColorScheme, "")
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteCustomXmlParts(ByVal iLevel As Integer)
    Dim oPart As Object

    iLevel = iLevel + 1

    On Error Resume Next

    If (m_oPresentation.CustomXMLParts.Count < 0) Then
    End If

    If (Err.Number <> 0) Then
        Err.Clear
        On Error GoTo 0
        Call WritePropValueNull(iLevel, "CustomXmlParts")
        Exit Sub
    End If

    On Error GoTo 0

    Call WriteBeginObjectList(iLevel, "CustomXmlParts")

    For Each oPart In m_oPresentation.CustomXMLParts
        Call WriteCustomXmlPart(iLevel, oPart)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteCustomXmlPart(ByVal iLevel As Integer, oPart As Object)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueBoolean(iLevel, "BuiltIn", oPart.BuiltIn)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "BuiltIn")
    End If

    Call WritePropValueString(iLevel, "Id", oPart.Id)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Id")
    End If

    Call WritePropValueString(iLevel, "NamespaceURI", oPart.NamespaceURI)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NamespaceURI")
    End If

    Call WritePropValueString(iLevel, "XML", oPart.XML)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "XML")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteDesigns(ByVal iLevel As Integer)
    Dim oDesign As PowerPoint.Design

    iLevel = iLevel + 1

    On Error Resume Next

    If (m_oPresentation.Designs.Count < 0) Then
    End If

    If (Err.Number <> 0) Then
        Err.Clear
        On Error GoTo 0
        Call WritePropValueNull(iLevel, "Designs")
        Exit Sub
    End If

    On Error GoTo 0

    Call WriteBeginObjectList(iLevel, "Designs")

    For Each oDesign In m_oPresentation.Designs
        Call WriteDesign(iLevel, oDesign)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteDesign(ByVal iLevel As Integer, oDesign As PowerPoint.Design)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", oDesign.Name)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Name")
    End If

    Call WritePropValueLong(iLevel, "Index", oDesign.Index)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Index")
    End If

    Call WritePropValueLong(iLevel, "HasTitleMaster", oDesign.HasTitleMaster)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HasTitleMaster")
    End If

    Call WritePropValueLong(iLevel, "Preserved", oDesign.Preserved)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Preserved")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteExtraColors(ByVal iLevel As Integer)
    Dim nIndex As Long
    Dim nCount As Long

    iLevel = iLevel + 1

    On Error Resume Next

    nCount = m_oPresentation.ExtraColors.Count

    If (Err.Number <> 0) Then
        Err.Clear
        On Error GoTo 0
        Call WritePropValueNull(iLevel, "ExtraColors")
        Exit Sub
    End If

    On Error GoTo 0

    If (nCount = 0) Then
        Call WriteBeginEndObjectList(iLevel, "ExtraColors")
        Exit Sub
    End If

    Call WriteBeginObjectList(iLevel, "ExtraColors")

    For nIndex = 1 To nCount
        Call WriteListValueLong(iLevel + 1, m_oPresentation.ExtraColors.Item(nIndex))
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteFonts(ByVal iLevel As Integer)
    Dim oFont As PowerPoint.Font

    iLevel = iLevel + 1

    On Error Resume Next

    If (m_oPresentation.Fonts.Count < 0) Then
    End If

    If (Err.Number <> 0) Then
        Err.Clear
        On Error GoTo 0
        Call WritePropValueNull(iLevel, "Fonts")
        Exit Sub
    End If

    On Error GoTo 0

    Call WriteBeginObjectList(iLevel, "Fonts")

    For Each oFont In m_oPresentation.Fonts
        Call WritePresentationFont(iLevel, oFont)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WritePresentationFont(ByVal iLevel As Integer, oFont As PowerPoint.Font)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", oFont.Name)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Name")
    End If

    Call WritePropValueLong(iLevel, "Embeddable", oFont.Embeddable)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Embeddable")
    End If

    Call WritePropValueLong(iLevel, "Embedded", oFont.Embedded)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Embedded")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteGuides(ByVal iLevel As Integer)
    Dim nIndex As Long
    Dim nCount As Long

    iLevel = iLevel + 1

    On Error Resume Next

    nCount = m_oPresentation.Guides.Count

    If (Err.Number <> 0) Then
        Err.Clear
        On Error GoTo 0
        Call WritePropValueNull(iLevel, "Guides")
        Exit Sub
    End If

    On Error GoTo 0

    If (nCount = 0) Then
        Call WriteBeginEndObjectList(iLevel, "Guides")
        Exit Sub
    End If

    Call WriteBeginObjectList(iLevel, "Guides")

    For nIndex = 1 To nCount
        Call WriteGuide(iLevel, m_oPresentation.Guides.Item(nIndex))
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteGuide(ByVal iLevel As Integer, oGuide As Object)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    Call WriteColorFormat(iLevel, oGuide.Color, "Color")
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "Color")
    End If

    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "Orientation", oGuide.Orientation)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Orientation")
    End If

    Call WritePropValueSingleString(iLevel, "Position", oGuide.Position)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Position")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteSectionProperties(ByVal iLevel As Integer, oSectionProperties As PowerPoint.SectionProperties)
    Dim nSectionIndex As Long

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "SectionProperties")

    iLevel = iLevel + 1

    If (oSectionProperties.Count > 0) Then
        For nSectionIndex = 1 To oSectionProperties.Count
            Call WriteBeginObject(iLevel, "")

            iLevel = iLevel + 1

            Call WritePropValueLong(iLevel, "Index", nSectionIndex)
            Call WritePropValueString(iLevel, "Name", oSectionProperties.Name(nSectionIndex))
            Call WritePropValueString(iLevel, "SectionID", oSectionProperties.SectionID(nSectionIndex))
            Call WritePropValueLong(iLevel, "SlidesCount", oSectionProperties.SlidesCount(nSectionIndex))

            iLevel = iLevel - 1

            Call WriteEndObject(iLevel)
        Next
    End If

    iLevel = iLevel - 1

    Call WriteEndObjectList(iLevel)
End Sub

' -- Slides ---------------------------------------------------------------------

Private Sub WriteSlides(ByVal iLevel As Integer, oSlides As PowerPoint.Slides)
    Dim oSlide As PowerPoint.Slide

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "Slides")

    For Each oSlide In oSlides
        Call WriteSlide(iLevel, oSlide)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteSlide(ByVal iLevel As Integer, oSlide As PowerPoint.Slide)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    On Error Resume Next
    Call WriteSlideShowTransition(iLevel, oSlide.SlideShowTransition)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "SlideShowTransition")
    End If
    On Error GoTo 0

    Call WriteColorScheme(iLevel, oSlide.ColorScheme, "ColorScheme")

    Call WriteThemeColorScheme(iLevel, oSlide.ThemeColorScheme)

    On Error Resume Next
    Call WriteHeadersFooters(iLevel, oSlide.HeadersFooters)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "HeadersFooters")
    End If
    On Error GoTo 0

    Call WriteSlideRange(iLevel, oSlide.NotesPage, "NotesPage")

    Call WriteShapeRange(iLevel, oSlide.Background, "Background")

    Call WriteShapes(iLevel, oSlide.Shapes)

    Call WriteTags(iLevel, oSlide.Tags)

    Call WriteComments(iLevel, oSlide.Comments, "Comments")

    Call WriteHyperlinks(iLevel, oSlide.Hyperlinks, "Hyperlinks")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", oSlide.Name)
    Call WritePropValueString(iLevel, "CustomLayout_Name", oSlide.CustomLayout.Name)

    On Error Resume Next
    Call WritePropValueString(iLevel, "Design_Name", oSlide.Design.Name)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Design_Name")
    End If
    On Error GoTo 0

    Call WritePropValueLong(iLevel, "SlideID", oSlide.SlideID)
    Call WritePropValueLong(iLevel, "SlideIndex", oSlide.SlideIndex)
    Call WritePropValueLong(iLevel, "SlideNumber", oSlide.SlideNumber)
    Call WritePropValueLong(iLevel, "sectionIndex", oSlide.sectionIndex)
    Call WritePropValueInteger(iLevel, "BackgroundStyle", oSlide.BackgroundStyle)
    Call WritePropValueInteger(iLevel, "DisplayMasterShapes", oSlide.DisplayMasterShapes)
    Call WritePropValueInteger(iLevel, "FollowMasterBackground", oSlide.FollowMasterBackground)
    Call WritePropValueInteger(iLevel, "HasNotesPage", oSlide.HasNotesPage)
    Call WritePropValueInteger(iLevel, "Layout", oSlide.Layout)
    Call WritePropValueLong(iLevel, "PrintSteps", oSlide.PrintSteps)

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteSlideRange(ByVal iLevel As Integer, oSlideRange As PowerPoint.SlideRange, ByVal sNameAs As String)
    iLevel = iLevel + 1

    If (sNameAs = "") Then
        sNameAs = "SlideRange"
    End If

    Call WriteBeginObject(iLevel, sNameAs)

    Call WriteShapes(iLevel, oSlideRange.Shapes)

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", oSlideRange.Name)

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteShapeRange(ByVal iLevel As Integer, oShapeRange As PowerPoint.ShapeRange, ByVal sNameAs As String)
    Dim nIndex As Long

    iLevel = iLevel + 1

    If (sNameAs = "") Then
        sNameAs = "ShapeRange"
    End If

    Call WriteBeginObjectList(iLevel, sNameAs)

    If (oShapeRange.Count > 0) Then
        For nIndex = 1 To oShapeRange.Count
            Call WriteShape(iLevel, oShapeRange(nIndex))
        Next
    End If

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteHeadersFooters(ByVal iLevel As Integer, oHeadersFooters As PowerPoint.HeadersFooters)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "HeadersFooters")

    ' A failing getter (e.g. Header on a slide, which only exists on notes and handout
    ' pages) must still write the property as null so the canonical schema is complete.
    Call WriteHeaderFooter(iLevel, oHeadersFooters.DateAndTime, "DateAndTime")
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "DateAndTime")
    End If

    Call WriteHeaderFooter(iLevel, oHeadersFooters.Footer, "Footer")
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "Footer")
    End If

    Call WriteHeaderFooter(iLevel, oHeadersFooters.Header, "Header")
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "Header")
    End If

    Call WriteHeaderFooter(iLevel, oHeadersFooters.SlideNumber, "SlideNumber")
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "SlideNumber")
    End If

    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "DisplayOnTitleSlide", oHeadersFooters.DisplayOnTitleSlide)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "DisplayOnTitleSlide")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteHeaderFooter(ByVal iLevel As Integer, oHeaderFooter As Object, sNameAs As String)
    On Error Resume Next

    If (oHeaderFooter Is Nothing) Then
        Call WritePropValueNull(iLevel + 1, sNameAs)
        Exit Sub
    End If

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, sNameAs)

    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "Format", oHeaderFooter.Format)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Format")
    End If

    Call WritePropValueString(iLevel, "Text", oHeaderFooter.Text)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Text")
    End If

    Call WritePropValueLong(iLevel, "UseFormat", oHeaderFooter.UseFormat)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "UseFormat")
    End If

    Call WritePropValueLong(iLevel, "Visible", oHeaderFooter.Visible)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Visible")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Color schemes ----------------------------------------------------------------

Private Sub WriteColorScheme(ByVal iLevel As Integer, oColorScheme As PowerPoint.ColorScheme, sNameAs As String)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, sNameAs)

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Accent1", oColorScheme.Colors(ppAccent1).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Accent1")
    End If

    Call WritePropValueString(iLevel, "Accent2", oColorScheme.Colors(ppAccent2).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Accent2")
    End If

    Call WritePropValueString(iLevel, "Accent3", oColorScheme.Colors(ppAccent3).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Accent3")
    End If

    Call WritePropValueString(iLevel, "Background", oColorScheme.Colors(ppBackground).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Background")
    End If

    Call WritePropValueString(iLevel, "Fill", oColorScheme.Colors(ppFill).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Fill")
    End If

    Call WritePropValueString(iLevel, "Foreground", oColorScheme.Colors(ppForeground).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Foreground")
    End If

    Call WritePropValueString(iLevel, "NotSchemeColor", oColorScheme.Colors(ppNotSchemeColor).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NotSchemeColor")
    End If

    Call WritePropValueString(iLevel, "SchemeColorMixed", oColorScheme.Colors(ppSchemeColorMixed).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "SchemeColorMixed")
    End If

    Call WritePropValueString(iLevel, "Shadow", oColorScheme.Colors(ppShadow).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Shadow")
    End If

    Call WritePropValueString(iLevel, "Title", oColorScheme.Colors(ppTitle).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Title")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteThemeColorScheme(ByVal iLevel As Integer, oThemeColorScheme As Office.ThemeColorScheme)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "ThemeColorScheme")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "ThemeAccent1", oThemeColorScheme.Colors(msoThemeAccent1).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ThemeAccent1")
    End If

    Call WritePropValueString(iLevel, "ThemeAccent2", oThemeColorScheme.Colors(msoThemeAccent2).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ThemeAccent2")
    End If

    Call WritePropValueString(iLevel, "ThemeAccent3", oThemeColorScheme.Colors(msoThemeAccent3).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ThemeAccent3")
    End If

    Call WritePropValueString(iLevel, "ThemeAccent4", oThemeColorScheme.Colors(msoThemeAccent4).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ThemeAccent4")
    End If

    Call WritePropValueString(iLevel, "ThemeAccent5", oThemeColorScheme.Colors(msoThemeAccent5).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ThemeAccent5")
    End If

    Call WritePropValueString(iLevel, "ThemeAccent6", oThemeColorScheme.Colors(msoThemeAccent6).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ThemeAccent6")
    End If

    Call WritePropValueString(iLevel, "ThemeDark1", oThemeColorScheme.Colors(msoThemeDark1).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ThemeDark1")
    End If

    Call WritePropValueString(iLevel, "ThemeDark2", oThemeColorScheme.Colors(msoThemeDark2).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ThemeDark2")
    End If

    Call WritePropValueString(iLevel, "ThemeFollowedHyperlink", oThemeColorScheme.Colors(msoThemeFollowedHyperlink).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ThemeFollowedHyperlink")
    End If

    Call WritePropValueString(iLevel, "ThemeHyperlink", oThemeColorScheme.Colors(msoThemeHyperlink).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ThemeHyperlink")
    End If

    Call WritePropValueString(iLevel, "ThemeLight1", oThemeColorScheme.Colors(msoThemeLight1).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ThemeLight1")
    End If

    Call WritePropValueString(iLevel, "ThemeLight2", oThemeColorScheme.Colors(msoThemeLight2).RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ThemeLight2")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteColorFormat(ByVal iLevel As Integer, oColorFormat As PowerPoint.ColorFormat, sNameAs As String)
    On Error Resume Next

    iLevel = iLevel + 1

    If (sNameAs = "") Then
        sNameAs = "ColorFormat"
    End If

    Call WriteBeginObject(iLevel, sNameAs)

    iLevel = iLevel + 1

    Call WritePropValueSingleString(iLevel, "Brightness", oColorFormat.Brightness)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Brightness")
    End If

    Call WritePropValueLong(iLevel, "Creator", oColorFormat.Creator)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Creator")
    End If

    Call WritePropValueLong(iLevel, "ObjectThemeColor", oColorFormat.ObjectThemeColor)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ObjectThemeColor")
    End If

    Call WritePropValueLong(iLevel, "RGB", oColorFormat.RGB)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "RGB")
    End If

    Call WritePropValueLong(iLevel, "SchemeColor", oColorFormat.SchemeColor)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "SchemeColor")
    End If

    Call WritePropValueSingleString(iLevel, "TintAndShade", oColorFormat.TintAndShade)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TintAndShade")
    End If

    Call WritePropValueLong(iLevel, "Type", oColorFormat.Type)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Type")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Shapes ---------------------------------------------------------------------

Private Sub WriteShapes(ByVal iLevel As Integer, ppShapes As PowerPoint.Shapes)
    Dim oShape As PowerPoint.Shape

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "Shapes")

    For Each oShape In ppShapes
        Call WriteShape(iLevel, oShape)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteGroupItems(ByVal iLevel As Integer, oGroupShapes As PowerPoint.GroupShapes)
    Dim oShape As PowerPoint.Shape

    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "GroupItems")

    For Each oShape In oGroupShapes
        Call WriteShape(iLevel, oShape)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteShape(ByVal iLevel As Integer, oShape As PowerPoint.Shape)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    Call WriteAnimationSettings(iLevel, oShape.AnimationSettings)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "AnimationSettings")
    End If

    Call WriteFillFormat(iLevel, oShape.Fill, "Fill")
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "Fill")
    End If

    Call WriteLineFormat(iLevel, oShape.Line, "Line")
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "Line")
    End If

    If (IsPictureShape(oShape) = True) Then
        Call WritePictureFormat(iLevel, oShape.PictureFormat)
        If (Err.Number <> 0) Then
            Err.Clear
            Call WritePropValueNull(iLevel + 1, "PictureFormat")
        End If
    Else
        Call WritePropValueNull(iLevel + 1, "PictureFormat")
    End If

    Call WritePlaceholderFormat(iLevel, oShape.PlaceholderFormat, "PlaceholderFormat")
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "PlaceholderFormat")
    End If

    Call WriteTextFrame(iLevel, oShape.TextFrame)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "TextFrame")
    End If

    Call WriteTextFrame2(iLevel, oShape.TextFrame2)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "TextFrame2")
    End If

    If (IsPictureShape(oShape) = True) Then
        Call WriteImageBlob(iLevel, oShape)
    Else
        Call WritePropValueNull(iLevel + 1, "Image")
    End If

    Call WriteActionSettings(iLevel, oShape.ActionSettings)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "ActionSettings")
    End If

    Call WriteTags(iLevel, oShape.Tags)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "Tags")
    End If

    If (oShape.Type = msoGroup) Then
        Call WriteGroupItems(iLevel, oShape.GroupItems)
    Else
        Call WritePropValueNull(iLevel + 1, "GroupItems")
    End If

    '** Shape Fields
    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "Id", oShape.Id)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Id")
    End If

    Call WritePropValueString(iLevel, "Name", oShape.Name)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Name")
    End If

    Call WritePropValueInteger(iLevel, "Type", oShape.Type)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Type")
    End If

    Call WritePropValueInteger(iLevel, "AutoShapeType", oShape.AutoShapeType)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AutoShapeType")
    End If

    Call WritePropValueString(iLevel, "AlternativeText", oShape.AlternativeText)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AlternativeText")
    End If

    Call WritePropValueString(iLevel, "Title", oShape.Title)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Title")
    End If

    Call WritePropValueSingleString(iLevel, "Left", oShape.Left)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Left")
    End If

    Call WritePropValueSingleString(iLevel, "Top", oShape.Top)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Top")
    End If

    Call WritePropValueSingleString(iLevel, "Width", oShape.Width)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Width")
    End If

    Call WritePropValueSingleString(iLevel, "Height", oShape.Height)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Height")
    End If

    Call WritePropValueSingleString(iLevel, "Rotation", oShape.Rotation)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Rotation")
    End If

    Call WritePropValueLong(iLevel, "ZOrderPosition", oShape.ZOrderPosition)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ZOrderPosition")
    End If

    Call WritePropValueInteger(iLevel, "Visible", oShape.Visible)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Visible")
    End If

    Call WritePropValueInteger(iLevel, "BackgroundStyle", oShape.BackgroundStyle)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "BackgroundStyle")
    End If

    Call WritePropValueInteger(iLevel, "BlackWhiteMode", oShape.BlackWhiteMode)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "BlackWhiteMode")
    End If

    Call WritePropValueInteger(iLevel, "Child", oShape.Child)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Child")
    End If

    Call WritePropValueLong(iLevel, "ConnectionSiteCount", oShape.ConnectionSiteCount)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ConnectionSiteCount")
    End If

    Call WritePropValueInteger(iLevel, "Connector", oShape.Connector)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Connector")
    End If

    Call WritePropValueLong(iLevel, "Creator", oShape.Creator)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Creator")
    End If

    Call WritePropValueInteger(iLevel, "Decorative", oShape.Decorative)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Decorative")
    End If

    Call WritePropValueInteger(iLevel, "GraphicStyle", oShape.GraphicStyle)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "GraphicStyle")
    End If

    Call WritePropValueInteger(iLevel, "HasChart", oShape.HasChart)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HasChart")
    End If

    Call WritePropValueInteger(iLevel, "HasInkXML", oShape.HasInkXML)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HasInkXML")
    End If

    Call WritePropValueInteger(iLevel, "HasSectionZoom", oShape.HasSectionZoom)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HasSectionZoom")
    End If

    Call WritePropValueInteger(iLevel, "HasSmartArt", oShape.HasSmartArt)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HasSmartArt")
    End If

    Call WritePropValueInteger(iLevel, "HasTable", oShape.HasTable)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HasTable")
    End If

    Call WritePropValueInteger(iLevel, "HasTextFrame", oShape.HasTextFrame)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HasTextFrame")
    End If

    Call WritePropValueInteger(iLevel, "HorizontalFlip", oShape.HorizontalFlip)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HorizontalFlip")
    End If

    Call WritePropValueString(iLevel, "InkXML", oShape.InkXML)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "InkXML")
    End If

    Call WritePropValueInteger(iLevel, "IsNarration", oShape.IsNarration)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "IsNarration")
    End If

    Call WritePropValueInteger(iLevel, "LockAspectRatio", oShape.LockAspectRatio)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "LockAspectRatio")
    End If

    Call WritePropValueInteger(iLevel, "ShapeStyle", oShape.ShapeStyle)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ShapeStyle")
    End If

    Call WritePropValueInteger(iLevel, "VerticalFlip", oShape.VerticalFlip)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "VerticalFlip")
    End If

    Call WritePropValueString(iLevel, "Vertices", oShape.Vertices)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Vertices")
    End If

    Call WritePropValueInteger(iLevel, "MediaType", oShape.MediaType)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "MediaType")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Function IsPictureShape(oShape As PowerPoint.Shape) As Boolean
    On Error Resume Next

    IsPictureShape = (oShape.Type = msoPicture) Or (oShape.Type = msoLinkedPicture)

    If (Err.Number <> 0) Then
        Err.Clear
        IsPictureShape = False
    End If

    On Error GoTo 0
End Function

' Writes the shape image using the writer blob options: base64 (default) or a
' separate file that the JSON references.
Private Sub WriteImageBlob(ByVal iLevel As Integer, oShape As PowerPoint.Shape)
    Dim oFSO As Scripting.FileSystemObject
    Dim sTempPath As String
    Dim sFileName As String

    iLevel = iLevel + 1

    Set oFSO = New Scripting.FileSystemObject

    sTempPath = oFSO.BuildPath(oFSO.GetSpecialFolder(2), oFSO.GetTempName() & ".png")

    On Error Resume Next
    oShape.Export sTempPath, ppShapeFormatPNG
    If (Err.Number <> 0) Then
        Err.Clear
        On Error GoTo 0
        Call WritePropValueNull(iLevel, "Image")
        Exit Sub
    End If
    On Error GoTo 0

    Call WriteBeginObject(iLevel, "Image")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Extension", "png")

    If (m_eBlobOutput = jsonBlobBase64) Then
        Call WritePropValueString(iLevel, "Base64", ReadFileAsBase64(sTempPath))
        Call WritePropValueNull(iLevel, "FileName")
    Else
        Call EnsureFolderExists(m_sBlobFolderPath)

        sFileName = CStr(oShape.Id) & ".png"
        oFSO.CopyFile sTempPath, oFSO.BuildPath(m_sBlobFolderPath, sFileName), True

        Call WritePropValueNull(iLevel, "Base64")
        Call WritePropValueString(iLevel, "FileName", m_sBlobFolderName & "/" & sFileName)
    End If

    On Error Resume Next
    oFSO.DeleteFile sTempPath
    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Shape formats ----------------------------------------------------------------

Private Sub WriteAnimationSettings(ByVal iLevel As Integer, oAnimationSettings As PowerPoint.AnimationSettings)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "AnimationSettings")

    Call WriteColorFormat(iLevel, oAnimationSettings.DimColor, "DimColor")
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "DimColor")
    End If

    Call WritePlaySettings(iLevel, oAnimationSettings.PlaySettings)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "PlaySettings")
    End If

    Call WriteSoundEffect(iLevel, oAnimationSettings.SoundEffect, "SoundEffect")
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "SoundEffect")
    End If

    iLevel = iLevel + 1

    Call WritePropValueInteger(iLevel, "AdvanceMode", oAnimationSettings.AdvanceMode)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AdvanceMode")
    End If

    Call WritePropValueSingleString(iLevel, "AdvanceTime", oAnimationSettings.AdvanceTime)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AdvanceTime")
    End If

    Call WritePropValueInteger(iLevel, "AfterEffect", oAnimationSettings.AfterEffect)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AfterEffect")
    End If

    Call WritePropValueInteger(iLevel, "Animate", oAnimationSettings.Animate)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Animate")
    End If

    Call WritePropValueInteger(iLevel, "AnimateBackground", oAnimationSettings.AnimateBackground)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AnimateBackground")
    End If

    Call WritePropValueInteger(iLevel, "AnimateTextInReverse", oAnimationSettings.AnimateTextInReverse)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AnimateTextInReverse")
    End If

    Call WritePropValueLong(iLevel, "AnimationOrder", oAnimationSettings.AnimationOrder)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AnimationOrder")
    End If

    Call WritePropValueInteger(iLevel, "EntryEffect", oAnimationSettings.EntryEffect)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "EntryEffect")
    End If

    Call WritePropValueInteger(iLevel, "TextLevelEffect", oAnimationSettings.TextLevelEffect)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TextLevelEffect")
    End If

    Call WritePropValueInteger(iLevel, "TextUnitEffect", oAnimationSettings.TextUnitEffect)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TextUnitEffect")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WritePlaySettings(ByVal iLevel As Integer, oPlaySettings As PowerPoint.PlaySettings)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "PlaySettings")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "ActionVerb", oPlaySettings.ActionVerb)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ActionVerb")
    End If

    Call WritePropValueLong(iLevel, "HideWhileNotPlaying", oPlaySettings.HideWhileNotPlaying)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "HideWhileNotPlaying")
    End If

    Call WritePropValueLong(iLevel, "LoopUntilStopped", oPlaySettings.LoopUntilStopped)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "LoopUntilStopped")
    End If

    Call WritePropValueLong(iLevel, "PauseAnimation", oPlaySettings.PauseAnimation)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "PauseAnimation")
    End If

    Call WritePropValueLong(iLevel, "PlayOnEntry", oPlaySettings.PlayOnEntry)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "PlayOnEntry")
    End If

    Call WritePropValueLong(iLevel, "RewindMovie", oPlaySettings.RewindMovie)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "RewindMovie")
    End If

    Call WritePropValueLong(iLevel, "StopAfterSlides", oPlaySettings.StopAfterSlides)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "StopAfterSlides")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteFillFormat(ByVal iLevel As Integer, oFillFormat As PowerPoint.FillFormat, sNameAs As String)
    On Error Resume Next

    iLevel = iLevel + 1

    If (sNameAs = "") Then
        sNameAs = "Fill"
    End If

    Call WriteBeginObject(iLevel, sNameAs)

    Call WriteColorFormat(iLevel, oFillFormat.BackColor, "BackColor")
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "BackColor")
    End If

    Call WriteColorFormat(iLevel, oFillFormat.ForeColor, "ForeColor")
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "ForeColor")
    End If

    iLevel = iLevel + 1

    Call WritePropValueSingleString(iLevel, "GradientAngle", oFillFormat.GradientAngle)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "GradientAngle")
    End If

    Call WritePropValueLong(iLevel, "GradientColorType", oFillFormat.GradientColorType)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "GradientColorType")
    End If

    Call WritePropValueSingleString(iLevel, "GradientDegree", oFillFormat.GradientDegree)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "GradientDegree")
    End If

    Call WritePropValueLong(iLevel, "GradientStyle", oFillFormat.GradientStyle)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "GradientStyle")
    End If

    Call WritePropValueLong(iLevel, "GradientVariant", oFillFormat.GradientVariant)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "GradientVariant")
    End If

    Call WritePropValueLong(iLevel, "Pattern", oFillFormat.Pattern)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Pattern")
    End If

    Call WritePropValueLong(iLevel, "RotateWithObject", oFillFormat.RotateWithObject)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "RotateWithObject")
    End If

    Call WritePropValueLong(iLevel, "TextureAlignment", oFillFormat.TextureAlignment)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TextureAlignment")
    End If

    Call WritePropValueSingleString(iLevel, "TextureHorizontalScale", oFillFormat.TextureHorizontalScale)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TextureHorizontalScale")
    End If

    Call WritePropValueString(iLevel, "TextureName", oFillFormat.TextureName)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TextureName")
    End If

    Call WritePropValueSingleString(iLevel, "TextureOffsetX", oFillFormat.TextureOffsetX)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TextureOffsetX")
    End If

    Call WritePropValueSingleString(iLevel, "TextureOffsetY", oFillFormat.TextureOffsetY)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TextureOffsetY")
    End If

    Call WritePropValueLong(iLevel, "TextureTile", oFillFormat.TextureTile)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TextureTile")
    End If

    Call WritePropValueLong(iLevel, "TextureType", oFillFormat.TextureType)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TextureType")
    End If

    Call WritePropValueSingleString(iLevel, "TextureVerticalScale", oFillFormat.TextureVerticalScale)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TextureVerticalScale")
    End If

    Call WritePropValueSingleString(iLevel, "Transparency", oFillFormat.Transparency)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Transparency")
    End If

    Call WritePropValueLong(iLevel, "Type", oFillFormat.Type)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Type")
    End If

    Call WritePropValueLong(iLevel, "Visible", oFillFormat.Visible)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Visible")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteLineFormat(ByVal iLevel As Integer, oLineFormat As PowerPoint.LineFormat, sNameAs As String)
    On Error Resume Next

    iLevel = iLevel + 1

    If (sNameAs = "") Then
        sNameAs = "LineFormat"
    End If

    Call WriteBeginObject(iLevel, sNameAs)

    Call WriteColorFormat(iLevel, oLineFormat.BackColor, "BackColor")
    Call WriteColorFormat(iLevel, oLineFormat.ForeColor, "ForeColor")

    iLevel = iLevel + 1

    Call WritePropValueInteger(iLevel, "BeginArrowheadLength", oLineFormat.BeginArrowheadLength)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "BeginArrowheadLength")
    End If

    Call WritePropValueInteger(iLevel, "BeginArrowheadStyle", oLineFormat.BeginArrowheadStyle)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "BeginArrowheadStyle")
    End If

    Call WritePropValueInteger(iLevel, "BeginArrowheadWidth", oLineFormat.BeginArrowheadWidth)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "BeginArrowheadWidth")
    End If

    Call WritePropValueLong(iLevel, "Creator", oLineFormat.Creator)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Creator")
    End If

    Call WritePropValueInteger(iLevel, "DashStyle", oLineFormat.DashStyle)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "DashStyle")
    End If

    Call WritePropValueInteger(iLevel, "EndArrowheadLength", oLineFormat.EndArrowheadLength)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "EndArrowheadLength")
    End If

    Call WritePropValueInteger(iLevel, "EndArrowheadStyle", oLineFormat.EndArrowheadStyle)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "EndArrowheadStyle")
    End If

    Call WritePropValueInteger(iLevel, "EndArrowheadWidth", oLineFormat.EndArrowheadWidth)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "EndArrowheadWidth")
    End If

    Call WritePropValueInteger(iLevel, "InsetPen", oLineFormat.InsetPen)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "InsetPen")
    End If

    Call WritePropValueInteger(iLevel, "Pattern", oLineFormat.Pattern)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Pattern")
    End If

    Call WritePropValueInteger(iLevel, "Style", oLineFormat.Style)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Style")
    End If

    Call WritePropValueSingleString(iLevel, "Transparency", oLineFormat.Transparency)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Transparency")
    End If

    Call WritePropValueInteger(iLevel, "Visible", oLineFormat.Visible)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Visible")
    End If

    Call WritePropValueSingleString(iLevel, "Weight", oLineFormat.Weight)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Weight")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WritePictureFormat(ByVal iLevel As Integer, oPictureFormat As PowerPoint.PictureFormat)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "PictureFormat")

    iLevel = iLevel + 1

    Call WritePropValueSingleString(iLevel, "Brightness", oPictureFormat.Brightness)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Brightness")
    End If

    Call WritePropValueLong(iLevel, "ColorType", oPictureFormat.ColorType)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ColorType")
    End If

    Call WritePropValueSingleString(iLevel, "Contrast", oPictureFormat.Contrast)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Contrast")
    End If

    Call WritePropValueSingleString(iLevel, "CropBottom", oPictureFormat.CropBottom)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "CropBottom")
    End If

    Call WritePropValueSingleString(iLevel, "CropLeft", oPictureFormat.CropLeft)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "CropLeft")
    End If

    Call WritePropValueSingleString(iLevel, "CropRight", oPictureFormat.CropRight)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "CropRight")
    End If

    Call WritePropValueSingleString(iLevel, "CropTop", oPictureFormat.CropTop)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "CropTop")
    End If

    Call WritePropValueLong(iLevel, "TransparencyColor", oPictureFormat.TransparencyColor)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TransparencyColor")
    End If

    Call WritePropValueLong(iLevel, "TransparentBackground", oPictureFormat.TransparentBackground)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TransparentBackground")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WritePlaceholderFormat(ByVal iLevel As Integer, oPlaceholderFormat As PowerPoint.PlaceholderFormat, sNameAs As String)
    On Error Resume Next

    iLevel = iLevel + 1

    If (sNameAs = "") Then
        sNameAs = "PlaceholderFormat"
    End If

    Call WriteBeginObject(iLevel, sNameAs)

    iLevel = iLevel + 1

    Call WritePropValueInteger(iLevel, "ContainedType", oPlaceholderFormat.ContainedType)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ContainedType")
    End If

    Call WritePropValueString(iLevel, "Name", oPlaceholderFormat.Name)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Name")
    End If

    Call WritePropValueLong(iLevel, "Position", oPlaceholderFormat.Position)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Position")
    End If

    Call WritePropValueInteger(iLevel, "Type", oPlaceholderFormat.Type)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Type")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Text -----------------------------------------------------------------------

Private Sub WriteTextFrame(ByVal iLevel As Integer, oTextFrame As PowerPoint.TextFrame)
    Dim oTextRange As TextRange

    iLevel = iLevel + 1

    On Error Resume Next
    Set oTextRange = oTextFrame.TextRange
    If (Err.Number <> 0) Then
        Call WritePropValueNull(iLevel, "TextFrame")
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Call WriteBeginObject(iLevel, "TextFrame")

    Call WriteTextRange(iLevel, oTextFrame.TextRange)

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteTextRange(ByVal iLevel As Integer, oTextRange As PowerPoint.TextRange)
    Dim sText As String

    iLevel = iLevel + 1

    On Error Resume Next
    sText = oTextRange.Text
    If (Err.Number <> 0) Then
        Call WritePropValueNull(iLevel, "TextRange")
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Call WriteBeginObject(iLevel, "TextRange")

    Call WriteFont(iLevel, oTextRange.Font, "Font")

    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "Length", oTextRange.Length)
    Call WritePropValueLong(iLevel, "Start", oTextRange.Start)
    Call WritePropValueString(iLevel, "Text", oTextRange.Text)

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteFont(ByVal iLevel As Integer, oFont As PowerPoint.Font, sNameAs As String)
    On Error Resume Next

    iLevel = iLevel + 1

    If (sNameAs = "") Then
        sNameAs = "Font"
    End If

    Call WriteBeginObject(iLevel, sNameAs)

    Call WriteColorFormat(iLevel, oFont.Color, "Color")

    iLevel = iLevel + 1

    Call WritePropValueInteger(iLevel, "AutoRotateNumbers", oFont.AutoRotateNumbers)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AutoRotateNumbers")
    End If

    Call WritePropValueSingleString(iLevel, "BaselineOffset", oFont.BaselineOffset)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "BaselineOffset")
    End If

    Call WritePropValueInteger(iLevel, "Bold", oFont.Bold)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Bold")
    End If

    Call WritePropValueInteger(iLevel, "Emboss", oFont.Emboss)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Emboss")
    End If

    Call WritePropValueInteger(iLevel, "Italic", oFont.Italic)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Italic")
    End If

    Call WritePropValueString(iLevel, "Name", oFont.Name)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Name")
    End If

    Call WritePropValueString(iLevel, "NameAscii", oFont.NameAscii)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NameAscii")
    End If

    Call WritePropValueString(iLevel, "NameComplexScript", oFont.NameComplexScript)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NameComplexScript")
    End If

    Call WritePropValueString(iLevel, "NameFarEast", oFont.NameFarEast)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NameFarEast")
    End If

    Call WritePropValueString(iLevel, "NameOther", oFont.NameOther)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NameOther")
    End If

    Call WritePropValueInteger(iLevel, "Shadow", oFont.Shadow)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Shadow")
    End If

    Call WritePropValueSingleString(iLevel, "Size", oFont.Size)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Size")
    End If

    Call WritePropValueInteger(iLevel, "Subscript", oFont.Subscript)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Subscript")
    End If

    Call WritePropValueInteger(iLevel, "Superscript", oFont.Superscript)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Superscript")
    End If

    Call WritePropValueInteger(iLevel, "Underline", oFont.Underline)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Underline")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteTextFrame2(ByVal iLevel As Integer, oTextFrame2 As PowerPoint.TextFrame2)
    Dim oTextRange As TextRange2

    iLevel = iLevel + 1

    On Error Resume Next
    Set oTextRange = oTextFrame2.TextRange
    If (Err.Number <> 0) Then
        Call WritePropValueNull(iLevel, "TextFrame2")
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Call WriteBeginObject(iLevel, "TextFrame2")

    Call WriteTextRange2(iLevel, oTextFrame2.TextRange)

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteTextRange2(ByVal iLevel As Integer, oTextRange2 As Office.TextRange2)
    Dim sText As String

    iLevel = iLevel + 1

    On Error Resume Next
    sText = oTextRange2.Text
    If (Err.Number <> 0) Then
        Call WritePropValueNull(iLevel, "TextRange2")
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Call WriteBeginObject(iLevel, "TextRange2")

    Call WriteFont2(iLevel, oTextRange2.Font, "Font")

    iLevel = iLevel + 1

    Call WritePropValueLong(iLevel, "Length", oTextRange2.Length)
    Call WritePropValueLong(iLevel, "Start", oTextRange2.Start)
    Call WritePropValueString(iLevel, "Text", oTextRange2.Text)

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteFont2(ByVal iLevel As Integer, oFont2 As Font2, sNameAs As String)
    On Error Resume Next

    iLevel = iLevel + 1

    If (sNameAs = "") Then
        sNameAs = "Font"
    End If

    Call WriteBeginObject(iLevel, sNameAs)

    iLevel = iLevel + 1

    Call WritePropValueInteger(iLevel, "Allcaps", oFont2.Allcaps)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Allcaps")
    End If

    Call WritePropValueInteger(iLevel, "AutoRotateNumbers", oFont2.AutoRotateNumbers)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AutoRotateNumbers")
    End If

    Call WritePropValueSingleString(iLevel, "BaselineOffset", oFont2.BaselineOffset)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "BaselineOffset")
    End If

    Call WritePropValueInteger(iLevel, "Bold", oFont2.Bold)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Bold")
    End If

    Call WritePropValueInteger(iLevel, "Caps", oFont2.Caps)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Caps")
    End If

    Call WritePropValueInteger(iLevel, "DoubleStrikeThrough", oFont2.DoubleStrikeThrough)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "DoubleStrikeThrough")
    End If

    Call WritePropValueInteger(iLevel, "Equalize", oFont2.Equalize)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Equalize")
    End If

    Call WritePropValueInteger(iLevel, "Italic", oFont2.Italic)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Italic")
    End If

    Call WritePropValueSingleString(iLevel, "Kerning", oFont2.Kerning)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Kerning")
    End If

    Call WritePropValueString(iLevel, "Name", oFont2.Name)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Name")
    End If

    Call WritePropValueString(iLevel, "NameAscii", oFont2.NameAscii)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NameAscii")
    End If

    Call WritePropValueString(iLevel, "NameComplexScript", oFont2.NameComplexScript)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NameComplexScript")
    End If

    Call WritePropValueString(iLevel, "NameFarEast", oFont2.NameFarEast)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NameFarEast")
    End If

    Call WritePropValueString(iLevel, "NameOther", oFont2.NameOther)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "NameOther")
    End If

    Call WritePropValueSingleString(iLevel, "Size", oFont2.Size)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Size")
    End If

    Call WritePropValueInteger(iLevel, "Smallcaps", oFont2.Smallcaps)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Smallcaps")
    End If

    Call WritePropValueInteger(iLevel, "SoftEdgeFormat", oFont2.SoftEdgeFormat)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "SoftEdgeFormat")
    End If

    Call WritePropValueSingleString(iLevel, "Spacing", oFont2.Spacing)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Spacing")
    End If

    Call WritePropValueInteger(iLevel, "Strike", oFont2.Strike)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Strike")
    End If

    Call WritePropValueInteger(iLevel, "StrikeThrough", oFont2.StrikeThrough)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "StrikeThrough")
    End If

    Call WritePropValueInteger(iLevel, "Subscript", oFont2.Subscript)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Subscript")
    End If

    Call WritePropValueInteger(iLevel, "Superscript", oFont2.Superscript)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Superscript")
    End If

    Call WritePropValueInteger(iLevel, "UnderlineStyle", oFont2.UnderlineStyle)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "UnderlineStyle")
    End If

    Call WritePropValueInteger(iLevel, "WordArtformat", oFont2.WordArtFormat)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "WordArtformat")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Action settings ---------------------------------------------------------------

Private Sub WriteActionSettings(ByVal iLevel As Integer, oActionSettings As PowerPoint.ActionSettings)
    iLevel = iLevel + 1

    Call WriteBeginObjectList(iLevel, "ActionSettings")

    Call WriteActionSetting(iLevel, oActionSettings.Item(ppMouseClick), ppMouseClick)
    Call WriteActionSetting(iLevel, oActionSettings.Item(ppMouseOver), ppMouseOver)

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteActionSetting(ByVal iLevel As Integer, oActionSetting As PowerPoint.ActionSetting, ByVal ePpMouseActivation As PowerPoint.PpMouseActivation)
    On Error Resume Next

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    Call WriteActionSettingHyperlink(iLevel, oActionSetting)

    Call WriteSoundEffect(iLevel, oActionSetting.SoundEffect, "SoundEffect")
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "SoundEffect")
    End If

    iLevel = iLevel + 1

    Call WritePropValueInteger(iLevel, "MouseActivation", CInt(ePpMouseActivation))

    Call WritePropValueInteger(iLevel, "Action", oActionSetting.Action)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Action")
    End If

    Call WritePropValueString(iLevel, "ActionVerb", oActionSetting.ActionVerb)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ActionVerb")
    End If

    Call WritePropValueInteger(iLevel, "AnimateAction", oActionSetting.AnimateAction)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AnimateAction")
    End If

    Call WritePropValueString(iLevel, "Run", oActionSetting.Run)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Run")
    End If

    Call WritePropValueInteger(iLevel, "ShowAndReturn", oActionSetting.ShowAndReturn)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ShowAndReturn")
    End If

    Call WritePropValueString(iLevel, "SlideShowName", oActionSetting.SlideShowName)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "SlideShowName")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' Writes the ActionSetting hyperlink only when the action is a hyperlink action;
' otherwise the property is null.
Private Sub WriteActionSettingHyperlink(ByVal iLevel As Integer, oActionSetting As PowerPoint.ActionSetting)
    Dim oHyperlink As PowerPoint.Hyperlink

    On Error Resume Next

    If (oActionSetting.Action = ppActionHyperlink) Then
        Set oHyperlink = oActionSetting.Hyperlink
    End If

    If (Err.Number <> 0) Or (oHyperlink Is Nothing) Then
        Err.Clear
        On Error GoTo 0
        Call WritePropValueNull(iLevel + 1, "Hyperlink")
        Exit Sub
    End If

    On Error GoTo 0

    Call WriteHyperlink(iLevel, oHyperlink, "Hyperlink")
End Sub

' -- Transitions and sounds ---------------------------------------------------------

' Every property is written with its own error check so a property that raises
' (e.g. on the handout master's transition) becomes null instead of aborting the
' sub mid-object, which would leave unbalanced braces in the output.
Private Sub WriteSlideShowTransition(ByVal iLevel As Integer, oSlideShowTransition As PowerPoint.SlideShowTransition)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "SlideShowTransition")

    On Error Resume Next

    Call WriteSoundEffect(iLevel, oSlideShowTransition.SoundEffect, "SoundEffect")
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel + 1, "SoundEffect")
    End If

    iLevel = iLevel + 1

    Call WritePropValueInteger(iLevel, "AdvanceOnClick", oSlideShowTransition.AdvanceOnClick)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AdvanceOnClick")
    End If

    Call WritePropValueInteger(iLevel, "AdvanceOnTime", oSlideShowTransition.AdvanceOnTime)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AdvanceOnTime")
    End If

    Call WritePropValueSingleString(iLevel, "AdvanceTime", oSlideShowTransition.AdvanceTime)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "AdvanceTime")
    End If

    Call WritePropValueSingleString(iLevel, "Duration", oSlideShowTransition.Duration)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Duration")
    End If

    Call WritePropValueInteger(iLevel, "EntryEffect", oSlideShowTransition.EntryEffect)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "EntryEffect")
    End If

    Call WritePropValueInteger(iLevel, "Hidden", oSlideShowTransition.Hidden)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Hidden")
    End If

    Call WritePropValueInteger(iLevel, "LoopSoundUntilNext", oSlideShowTransition.LoopSoundUntilNext)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "LoopSoundUntilNext")
    End If

    Call WritePropValueInteger(iLevel, "Speed", oSlideShowTransition.Speed)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "Speed")
    End If

    On Error GoTo 0

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteSoundEffect(ByVal iLevel As Integer, oSoundEffect As PowerPoint.SoundEffect, sNameAs As String)
    Dim sName As String

    iLevel = iLevel + 1

    If (sNameAs = "") Then
        sNameAs = "SoundEffect"
    End If

    On Error Resume Next
    sName = oSoundEffect.Name
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, sNameAs)
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Call WriteBeginObject(iLevel, sNameAs)

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", oSoundEffect.Name)
    Call WritePropValueInteger(iLevel, "Type", oSoundEffect.Type)

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Comments, hyperlinks, tags ------------------------------------------------------

Private Sub WriteComments(ByVal iLevel As Integer, oComments As PowerPoint.Comments, sNameAs As String)
    Dim nIndex As Long

    iLevel = iLevel + 1

    If (sNameAs = "") Then
        sNameAs = "Comments"
    End If

    If (oComments.Count = 0) Then
        Call WritePropValueNull(iLevel, sNameAs)
        Exit Sub
    End If

    Call WriteBeginObjectList(iLevel, sNameAs)

    For nIndex = 1 To oComments.Count
        Call WriteComment(iLevel, oComments.Item(nIndex))
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteComment(ByVal iLevel As Integer, oComment As PowerPoint.Comment)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Author", oComment.Author)
    Call WritePropValueString(iLevel, "AuthorInitials", oComment.AuthorInitials)
    Call WritePropValueBoolean(iLevel, "Collapsed", oComment.Collapsed)
    Call WritePropValueDate(iLevel, "DateTime", oComment.DateTime)
    Call WritePropValueSingleString(iLevel, "Left", oComment.Left)
    Call WritePropValueString(iLevel, "ProviderID", oComment.ProviderID)
    Call WritePropValueString(iLevel, "Text", oComment.Text)
    Call WritePropValueSingleString(iLevel, "Top", oComment.Top)
    Call WritePropValueString(iLevel, "UserID", oComment.UserID)

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteHyperlinks(ByVal iLevel As Integer, oHyperlinks As PowerPoint.Hyperlinks, sNameAs As String)
    Dim nIndex As Long

    iLevel = iLevel + 1

    If (sNameAs = "") Then
        sNameAs = "Hyperlinks"
    End If

    If (oHyperlinks.Count = 0) Then
        Call WritePropValueNull(iLevel, sNameAs)
        Exit Sub
    End If

    Call WriteBeginObjectList(iLevel, sNameAs)

    For nIndex = 1 To oHyperlinks.Count
        Call WriteHyperlink(iLevel, oHyperlinks.Item(nIndex), "")
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteHyperlink(ByVal iLevel As Integer, oHyperlink As PowerPoint.Hyperlink, sNameAs As String)
    Dim sShapeName As String
    Dim sText As String

    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, sNameAs)

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Address", oHyperlink.Address)
    Call WritePropValueString(iLevel, "EmailSubject", oHyperlink.EmailSubject)
    Call WritePropValueString(iLevel, "ScreenTip", oHyperlink.ScreenTip)

    On Error Resume Next
    Call WritePropValueInteger(iLevel, "ShowAndReturn", oHyperlink.ShowAndReturn)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "ShowAndReturn")
    End If
    On Error GoTo 0

    Call WritePropValueString(iLevel, "SubAddress", oHyperlink.SubAddress)

    On Error Resume Next
    Call WritePropValueString(iLevel, "TextToDisplay", oHyperlink.TextToDisplay)
    If (Err.Number <> 0) Then
        Err.Clear
        Call WritePropValueNull(iLevel, "TextToDisplay")
    End If
    On Error GoTo 0

    Call WritePropValueInteger(iLevel, "Type", oHyperlink.Type)

    On Error Resume Next
    If (oHyperlink.Type = msoHyperlinkShape) Then
        sShapeName = oHyperlink.Parent.Parent.Name
    ElseIf (oHyperlink.Type = msoHyperlinkRange) Then
        sText = oHyperlink.Parent.Parent.Text
        sShapeName = oHyperlink.Parent.Parent.Parent.Parent.Name
    End If
    Err.Clear
    On Error GoTo 0

    Call WritePropValueString(iLevel, "ShapeName", sShapeName)
    Call WritePropValueString(iLevel, "Text", sText)

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

Private Sub WriteTags(ByVal iLevel As Integer, oTags As PowerPoint.Tags)
    Dim i As Integer

    iLevel = iLevel + 1

    If (oTags.Count = 0) Then
        Call WriteBeginEndObjectList(iLevel, "Tags")
        Exit Sub
    End If

    Call WriteBeginObjectList(iLevel, "Tags")

    For i = 1 To oTags.Count
        Call WriteTag(iLevel, oTags, i)
    Next

    Call WriteEndObjectList(iLevel)
End Sub

Private Sub WriteTag(ByVal iLevel As Integer, oTags As PowerPoint.Tags, iIndex As Integer)
    iLevel = iLevel + 1

    Call WriteBeginObject(iLevel, "")

    iLevel = iLevel + 1

    Call WritePropValueString(iLevel, "Name", oTags.Name(iIndex))
    Call WritePropValueString(iLevel, "Value", oTags.Value(iIndex))

    iLevel = iLevel - 1

    Call WriteEndObject(iLevel)
End Sub

' -- Utilities ----------------------------------------------------------------

Private Sub LogStatus(sStatus As String)
    Debug.Print sStatus
End Sub
