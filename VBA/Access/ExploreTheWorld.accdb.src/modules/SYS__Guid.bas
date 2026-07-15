Attribute VB_Name = "SYS__Guid"
Option Explicit

Private Type GUID_TYPE
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (GUID As GUID_TYPE) As LongPtr
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (GUID As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As LongPtr

Function NewAccessGuid()
    NewAccessGuid = GUIDFromString(NewGuid2())
End Function

Function NewAccessGuidString() As String
    NewAccessGuidString = "{guid " & NewGuid() & "}"
End Function

Function AccessGuidString(sGuid As String) As String
    If (Left(sGuid, 1) = "{") Then
        AccessGuidString = "{guid " & sGuid & "}"
    Else
        AccessGuidString = "{guid {" & sGuid & "}}"
    End If
End Function

Function NewGuid() As String
    On Error GoTo Error_NewGuid

    Dim GUID As GUID_TYPE
    Dim strGUID As String
    Dim retValue As LongPtr

    Const guidLength As Long = 39 'registry GUID format with null terminator {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}

    retValue = CoCreateGuid(GUID)
    If retValue = 0 Then
        strGUID = String$(guidLength, vbNullChar)
        retValue = StringFromGUID2(GUID, StrPtr(strGUID), guidLength)
        If retValue = guidLength Then
            ' valid GUID as a string
            NewGuid = "{" & Mid$(strGUID, 2, 36) & "}" ' removes the braces from the output
        End If
    End If
Exit_NewGuid:
    Exit Function
Error_NewGuid:
    MsgBox "Error: " & Err.Number & ". " & Err.Description
    Resume Exit_NewGuid
End Function

Public Function NewGuid2() As String
    Dim udtGUID As GUID_TYPE

    If (CoCreateGuid(udtGUID) = 0) Then
        NewGuid2 = _
        String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
        String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
        String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
        IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
        IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
        IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
        IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
        IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
        IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
        IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
        IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If
End Function

Function NewGuidAsByteArray() As Byte()
    Dim strGUID As String

    Dim i As Integer
    Dim j As Integer
    Dim sPos As Integer
    Dim OffSet As Integer
    Dim sGuid(0 To 2) As Byte
    Dim bytArray() As Byte

    ReDim bytArray(0 To 15) As Byte

    strGUID = NewGuid()

    sGuid(0) = 7
    sGuid(1) = 11
    sGuid(2) = 15

    OffSet = 0
    sPos = 0

    'AABBCCDD-AABB-CCDD-XXXX-XXXXXXXXXXXX 'Microsoft Access view.
    'DDCCBBAA-BBAA-DDCC-XXXX-XXXXXXXXXXXX 'SQLServer view.
    'Need to loop through to build the GUID byte array in the Microsoft
    'Access storage format since the first eight bytes are reversed.
    For i = 0 To UBound(sGuid)
        For j = sGuid(i) To (OffSet + 1) Step -2
            bytArray(sPos) = "&H" & Mid$(strGUID, j, 2)
            sPos = sPos + 1
        Next j
        OffSet = sGuid(i)
    Next i

    For i = 17 To 31 Step 2
        bytArray(sPos) = "&H" & Mid$(strGUID, i, 2)
        sPos = sPos + 1
    Next i

    NewGuidAsByteArray = bytArray()

End Function

Function GetCleanGuid(sGuid As String) As String
    Dim sCleanGuid As String

    sCleanGuid = Replace(sGuid, "{guid ", "")
    sCleanGuid = Replace(sCleanGuid, "}}", "}")

    GetCleanGuid = sCleanGuid
End Function
