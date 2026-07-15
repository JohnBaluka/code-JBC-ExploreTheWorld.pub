Attribute VB_Name = "MSO_JsonWriterCore"
Option Explicit

' Shared JSON emitter used by MSO_MsPowerPoint_JsonWriter, MSO_MsExcel_JsonWriter and
' MSO_MsWord_JsonWriter. Produces strictly valid JSON that is byte-identical to the
' output of JBC.ExploreTheWorld.DL.MsOffice.MsOfficeJsonSerializer (System.Text.Json):
'   - 2-space indent, CRLF line endings, UTF-8 without BOM, trailing newline
'   - no trailing commas (a pending-line buffer decides commas when the next line arrives)
'   - string escaping via json_Encode (JsonConverter.bas), aligned with the
'     System.Text.Json default encoder

Public Enum JsonBlobOutput
    jsonBlobBase64 = 0
    jsonBlobSeparateFiles = 1
End Enum

Private Const INDENT As Integer = 2

Private m_oTextStream As Object
Private m_sPendingLine As String
Private m_bPendingIsValue As Boolean
Private m_bHasPending As Boolean

Public Sub JsonWriter_Begin()
    m_sPendingLine = ""
    m_bPendingIsValue = False
    m_bHasPending = False

    Set m_oTextStream = CreateObject("ADODB.Stream")
    m_oTextStream.Type = 2 ' adTypeText
    m_oTextStream.Charset = "utf-8"
    m_oTextStream.Open

    Call EmitLine("{", False, False)
End Sub

Public Sub JsonWriter_End(sOutputFilePath As String)
    Call EmitLine("}", True, True)

    ' Flush the final pending line ("}") without a comma.
    If (m_bHasPending = True) Then
        m_oTextStream.WriteText m_sPendingLine & vbCrLf
        m_bHasPending = False
    End If

    Call SaveStreamWithoutBom(sOutputFilePath)

    Set m_oTextStream = Nothing
End Sub

' -- Structure ------------------------------------------------------------------

Public Sub WriteBeginObject(ByVal iLevel As Integer, sPropertyName As String)
    Dim sSpace As String

    sSpace = Space(iLevel * INDENT)

    If (sPropertyName = "") Then
        Call EmitLine(sSpace & "{", False, False)
    Else
        Call EmitLine(sSpace & Chr(34) & sPropertyName & Chr(34) & ": {", False, False)
    End If
End Sub

Public Sub WriteEndObject(ByVal iLevel As Integer)
    Dim sSpace As String

    sSpace = Space(iLevel * INDENT)

    Call EmitLine(sSpace & "}", True, True)
End Sub

Public Sub WriteBeginObjectList(ByVal iLevel As Integer, sPropertyName As String)
    Dim sSpace As String

    sSpace = Space(iLevel * INDENT)

    Call EmitLine(sSpace & Chr(34) & sPropertyName & Chr(34) & ": [", False, False)
End Sub

Public Sub WriteEndObjectList(ByVal iLevel As Integer)
    Dim sSpace As String

    sSpace = Space(iLevel * INDENT)

    Call EmitLine(sSpace & "]", True, True)
End Sub

Public Sub WriteBeginEndObjectList(ByVal iLevel As Integer, sPropertyName As String)
    Dim sSpace As String

    sSpace = Space(iLevel * INDENT)

    Call EmitLine(sSpace & Chr(34) & sPropertyName & Chr(34) & ": []", True, False)
End Sub

' -- Property values ------------------------------------------------------------

Public Sub WritePropValueBoolean(ByVal iLevel As Integer, sPropertyName As String, bValue As Boolean)
    Dim sValue As String

    If (bValue = True) Then
        sValue = "true"
    Else
        sValue = "false"
    End If

    Call WriteRawValue(iLevel, sPropertyName, sValue)
End Sub

Public Sub WritePropValueInteger(ByVal iLevel As Integer, sPropertyName As String, iValue As Integer)
    Call WriteRawValue(iLevel, sPropertyName, CStr(iValue))
End Sub

Public Sub WritePropValueLong(ByVal iLevel As Integer, sPropertyName As String, lValue As Long)
    Call WriteRawValue(iLevel, sPropertyName, CStr(lValue))
End Sub

Public Sub WritePropValueSingleString(ByVal iLevel As Integer, sPropertyName As String, dValue As Single)
    Call WritePropValueString(iLevel, sPropertyName, CStr(dValue))
End Sub

Public Sub WritePropValueString(ByVal iLevel As Integer, sPropertyName As String, sValue As String)
    Call WriteRawValue(iLevel, sPropertyName, Chr(34) & json_Encode(sValue) & Chr(34))
End Sub

Public Sub WritePropValueDate(ByVal iLevel As Integer, sPropertyName As String, dtValue As Date)
    Call WritePropValueString(iLevel, sPropertyName, Format(dtValue, "mm/dd/yyyy hh:nn:ss AM/PM"))
End Sub

Public Sub WritePropValueNull(ByVal iLevel As Integer, sPropertyName As String)
    Call WriteRawValue(iLevel, sPropertyName, "null")
End Sub

' Writes a bare list item value (e.g. a number inside a JSON array).
Public Sub WriteListValueLong(ByVal iLevel As Integer, lValue As Long)
    Dim sSpace As String

    sSpace = Space(iLevel * INDENT)

    Call EmitLine(sSpace & CStr(lValue), True, False)
End Sub

Private Sub WriteRawValue(ByVal iLevel As Integer, sPropertyName As String, sRawValue As String)
    Dim sSpace As String

    sSpace = Space(iLevel * INDENT)

    Call EmitLine(sSpace & Chr(34) & sPropertyName & Chr(34) & ": " & sRawValue, True, False)
End Sub

' -- Emitter --------------------------------------------------------------------

' Buffers one line so the comma decision can be made when the next line arrives:
' a pending line that completes a value gets a comma unless the next line closes
' the containing object/list.
Private Sub EmitLine(sLine As String, bIsValueEnd As Boolean, bIsCloser As Boolean)
    Dim sCloser As String

    ' Collapse an empty list/object onto its opening line ('"X": []' / '"X": {}'),
    ' matching the System.Text.Json indented serializer.
    If (m_bHasPending = True) And (bIsCloser = True) Then
        sCloser = Trim$(sLine)
        If ((sCloser = "]") And (Right$(m_sPendingLine, 1) = "[")) Or _
           ((sCloser = "}") And (Right$(m_sPendingLine, 1) = "{")) Then
            m_sPendingLine = m_sPendingLine & sCloser
            m_bPendingIsValue = True
            Exit Sub
        End If
    End If

    If (m_bHasPending = True) Then
        If (m_bPendingIsValue = True) And (bIsCloser = False) Then
            m_oTextStream.WriteText m_sPendingLine & "," & vbCrLf
        Else
            m_oTextStream.WriteText m_sPendingLine & vbCrLf
        End If
    End If

    m_sPendingLine = sLine
    m_bPendingIsValue = bIsValueEnd
    m_bHasPending = True
End Sub

' -- File output ----------------------------------------------------------------

' Saves the open text stream as UTF-8 without BOM (matching MsOfficeJsonSerializer.WriteToFile).
Private Sub SaveStreamWithoutBom(sOutputFilePath As String)
    Dim oBinaryStream As Object
    Dim oFSO As Object

    Set oFSO = CreateObject("Scripting.FileSystemObject")

    If (oFSO.FileExists(sOutputFilePath) = True) Then
        oFSO.DeleteFile sOutputFilePath
    End If

    ' Strip the 3-byte BOM by copying from position 3 into a binary stream.
    m_oTextStream.Position = 0
    m_oTextStream.Type = 1 ' adTypeBinary
    m_oTextStream.Position = 3

    Set oBinaryStream = CreateObject("ADODB.Stream")
    oBinaryStream.Type = 1 ' adTypeBinary
    oBinaryStream.Open

    m_oTextStream.CopyTo oBinaryStream
    oBinaryStream.SaveToFile sOutputFilePath, 2 ' adSaveCreateOverWrite

    oBinaryStream.Close
    m_oTextStream.Close
End Sub

' -- Blob helpers ----------------------------------------------------------------

Public Function ReadFileAsBase64(sFilePath As String) As String
    Dim oStream As Object
    Dim oXml As Object
    Dim oNode As Object
    Dim sBase64 As String

    Set oStream = CreateObject("ADODB.Stream")
    oStream.Type = 1 ' adTypeBinary
    oStream.Open
    oStream.LoadFromFile sFilePath

    Set oXml = CreateObject("MSXML2.DOMDocument")
    Set oNode = oXml.createElement("b64")
    oNode.DataType = "bin.base64"
    oNode.nodeTypedValue = oStream.Read

    oStream.Close

    ' MSXML inserts line feeds every 76 characters; the canonical output has none.
    sBase64 = oNode.Text
    sBase64 = Replace(sBase64, vbCrLf, "")
    sBase64 = Replace(sBase64, vbLf, "")
    sBase64 = Replace(sBase64, vbCr, "")

    ReadFileAsBase64 = sBase64
End Function

Public Function GetBlobFolderPath(sOutputJsonFilePath As String, sBlobFolderPath As String) As String
    Dim oFSO As Object
    Dim sFolder As String
    Dim sName As String

    If (sBlobFolderPath <> "") Then
        GetBlobFolderPath = sBlobFolderPath
        Exit Function
    End If

    Set oFSO = CreateObject("Scripting.FileSystemObject")

    sFolder = oFSO.GetParentFolderName(sOutputJsonFilePath)
    sName = oFSO.GetBaseName(sOutputJsonFilePath) & "_Files"

    GetBlobFolderPath = oFSO.BuildPath(sFolder, sName)
End Function

Public Sub EnsureFolderExists(sFolderPath As String)
    Dim oFSO As Object

    Set oFSO = CreateObject("Scripting.FileSystemObject")

    If (oFSO.FolderExists(sFolderPath) = False) Then
        oFSO.CreateFolder sFolderPath
    End If
End Sub
