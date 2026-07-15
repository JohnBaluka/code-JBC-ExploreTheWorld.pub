Attribute VB_Name = "SYS__Strings"
Option Explicit

Function DQ(s As String)
    DQ = Chr(34) & s & Chr(34)
End Function

Function DDQ(s As String)
    DDQ = Chr(34) & Replace(s, Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

Function SQ(s As String)
    SQ = "'" & s & "'"
End Function

Function SSQ(s As String)
    SSQ = "'" & Replace(s, "'", "''") & "'"
End Function

Function DblQuote(s As String)
    DblQuote = Chr(34) & s & Chr(34)
End Function

Function DblDblQuote(s As String)
    DblDblQuote = Chr(34) & Replace(s, Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

Function SingleQuote(s As String)
    SingleQuote = "'" & s & "'"
End Function

Function SingleSingleQuote(s As String)
    SingleSingleQuote = "'" & Replace(s, "'", "''") & "'"
End Function

Function GetValue(vValue As Variant, oType As DataTypeEnum) As Variant
    Dim vRetValue As Variant
    Dim sValue As String

    sValue = Trim("" & vValue)

    If (sValue = "") Then
        Select Case oType
            Case dbText
                vRetValue = sValue
            Case Else
                vRetValue = vbNull
        End Select
    Else
        vRetValue = sValue
    End If

    GetValue = vRetValue
End Function

Function ReplaceInLine(sLine As String, vValue As Variant, oType As DataTypeEnum, sFormat As String, iStartPos As Integer, iStopPos As Integer) As String
    Dim vRetValue As Variant
    Dim sValue As String
    Dim sUpdatedLine As String
    Dim iLength As Integer
    Dim iMaxLength As Integer

    iMaxLength = iStopPos - iStartPos + 1

    sValue = Trim("" & vValue)

    If (sValue = "") Then
        sUpdatedLine = sLine & String(iMaxLength, " ")
    Else
        If (sFormat <> "") Then
            sValue = Format(sValue, sFormat)
        End If

        Select Case oType
            Case dbText
                'Do Nothing, for now
            Case dbDate
                'Do Nothing, for now
            Case Else
                sValue = Replace(sValue, ".", "")
        End Select

        iLength = Len(sValue)

        If (iLength > iMaxLength) Then
            Stop
        ElseIf iLength = iMaxLength Then
            sUpdatedLine = sLine & sValue
        Else
            sUpdatedLine = sLine & sValue & String(iMaxLength - iLength, " ")
        End If

    End If

'    Debug.Print sLine
'    Debug.Print sUpdatedLine
'    Debug.Print GetRulerLine1()
'    Debug.Print GetRulerLine2()

    ReplaceInLine = sUpdatedLine
End Function

Function GetRulerLine1() As String
    Dim sLine As String
    Dim sValue As String
    Dim i As Integer

    sValue = "1234567890"

    For i = 1 To 53
        sLine = sLine & sValue
    Next i

    GetRulerLine1 = sLine
End Function


Function GetRulerLine2() As String
    Dim sLine As String
    Dim sValue As String
    Dim i As Integer

    sValue = "1234567890"
    sValue = "        "

    For i = 1 To 53
        If (i > 9) Then
            sLine = sLine & sValue & Format(i, "00")
        Else
            sLine = sLine & sValue & " " & Format(i, "0")
        End If
    Next i

    GetRulerLine2 = sLine
End Function
