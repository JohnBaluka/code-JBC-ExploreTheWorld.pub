Attribute VB_Name = "ETW__FormManager"
Option Explicit

Global g_OpenForms As New Collection

Sub OpenForm(sName As String, sKeyFieldName As String, sKeyValue As String)

    If g_OpenForms Is Nothing Then
        Set g_OpenForms = New Collection
    End If

    Dim bFoundForm As Boolean
    bFoundForm = FindForm(sKeyValue)
    If bFoundForm Then
        g_OpenForms.Remove (sKeyValue)
    End If

    Dim sCaption As String
    Dim frmInfo As New APP__FormInfo
    frmInfo.Key = sKeyValue

    Select Case sName
        Case "cns_Country"
            sCaption = GetCnsCountryCaption(sKeyValue)
            Set frmInfo.Form = New Form_cns_Country
        Case Else
            Exit Sub
    End Select

    frmInfo.Caption = sCaption
    frmInfo.Form.Caption = sCaption
    frmInfo.Form.Filter = sKeyFieldName & "=" & sKeyValue
    frmInfo.Form.FilterOnLoad = True
    frmInfo.Form.FilterOn = True
    frmInfo.Form.Visible = True

    Call g_OpenForms.Add(frmInfo, sKeyValue)
End Sub

Function FindForm(sKeyValue As String) As Boolean
    FindForm = False

    If g_OpenForms Is Nothing Then
        Set g_OpenForms = New Collection
        Exit Function
    End If

    If g_OpenForms.Count = 0 Then
        Exit Function
    End If

    Dim frmInfo As APP__FormInfo
    For Each frmInfo In g_OpenForms
        If frmInfo.Key = sKeyValue Then
           FindForm = True
        End If
    Next
End Function

Function GotoForm(sKeyValue As String) As Boolean
    GotoForm = False

    If g_OpenForms Is Nothing Then
        Set g_OpenForms = New Collection
        Exit Function
    End If

    If g_OpenForms.Count = 0 Then
        Exit Function
    End If

    Dim f As Form

    Dim frmInfo As APP__FormInfo
    For Each frmInfo In g_OpenForms
        If frmInfo.Key = sKeyValue Then
            For Each f In Forms
                If (f.Caption = frmInfo.Caption) Then
                    f.SetFocus
                    GotoForm = True
                    Exit Function
                End If
            Next
        End If
    Next
End Function

Function GetCnsCountryCaption(sKeyValue As String) As String
    Dim sCaption As String
    sCaption = DLookup("Iso2", "cns_Country", "GUID = " & sKeyValue)
    GetCnsCountryCaption = sCaption
End Function

Function GetSourceCaption(sKeyValue As String) As String
    Dim sCaption As String
    sCaption = DLookup("Name", "WEB__Source", "GUID = " & sKeyValue)
    GetSourceCaption = sCaption
End Function
