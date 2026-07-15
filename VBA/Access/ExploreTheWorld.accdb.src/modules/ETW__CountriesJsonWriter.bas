Attribute VB_Name = "ETW__CountriesJsonWriter"
Option Explicit

Private m_oAdoStream As ADODB.Stream
Private m_sJson As String
Private m_iIndent As Integer

' ============================================================
' Public Entry Point
' ============================================================

Public Sub WriteCountriesToJsonFile(sOutputFilePath As String)
    m_sJson = ""
    m_iIndent = 0

    WriteBeginObject
        WritePropKey "rcc_Countries"
        WriteRccCountries
        WriteComma
        WritePropKey "cns_Countries"
        WriteCnsCountries
    WriteEndObject

    mFile.WriteAllText sOutputFilePath, m_sJson

    Debug.Print "WriteCountriesToJsonFile: Done - " & sOutputFilePath
End Sub

' ============================================================
' RCC Countries
' ============================================================

Private Sub WriteRccCountries()
    Dim db As DAO.Database
    Dim rsCountry As DAO.Recordset
    Dim rsName As DAO.Recordset
    Dim rsFlag As DAO.Recordset
    Dim rsCapitals As DAO.Recordset
    Dim rsLanguages As DAO.Recordset
    Dim rsCurrencies As DAO.Recordset
    Dim sCca2 As String
    Dim bFirstCountry As Boolean

    Set db = CurrentDb()
    Set rsCountry = db.OpenRecordset("SELECT * FROM rcc_Country ORDER BY Cca2", dbOpenSnapshot)

    WriteBeginArray
    bFirstCountry = True
    Do While Not rsCountry.EOF
        If Not bFirstCountry Then WriteComma
        bFirstCountry = False

        sCca2 = "" & rsCountry.Fields("Cca2").Value

        WriteBeginObject
            WritePropValueString "Cca2", sCca2
            WriteComma
            WritePropValueString "Cca3", Nz(rsCountry.Fields("Cca3").Value, "")
            WriteComma
            WritePropValueString "Region", Nz(rsCountry.Fields("Region").Value, "")
            WriteComma
            WritePropValueString "Subregion", Nz(rsCountry.Fields("Subregion").Value, "")
            WriteComma
            WritePropValueLong "Population", CLng(Nz(rsCountry.Fields("Population").Value, 0))

            ' Name
            Set rsName = db.OpenRecordset("SELECT * FROM rcc_CountryName WHERE Cca2 = " & SQ(sCca2), dbOpenSnapshot)
            If Not rsName.EOF Then
                WriteComma
                WritePropKey "Name"
                WriteBeginObject
                    WritePropValueString "Common", Nz(rsName.Fields("Common").Value, "")
                    WriteComma
                    WritePropValueString "Official", Nz(rsName.Fields("Official").Value, "")
                WriteEndObject
            End If
            rsName.Close

            ' Flag
            Set rsFlag = db.OpenRecordset("SELECT * FROM rcc_CountryFlag WHERE Cca2 = " & SQ(sCca2), dbOpenSnapshot)
            If Not rsFlag.EOF Then
                WriteComma
                WritePropKey "Flag"
                WriteBeginObject
                    WritePropValueString "Png", Nz(rsFlag.Fields("Png").Value, "")
                    WriteComma
                    WritePropValueString "Svg", Nz(rsFlag.Fields("Svg").Value, "")
                    WriteComma
                    WritePropValueString "Alt", Nz(rsFlag.Fields("Alt").Value, "")
                WriteEndObject
            End If
            rsFlag.Close

            ' Capitals
            WriteComma
            WritePropKey "Capitals"
            Set rsCapitals = db.OpenRecordset("SELECT * FROM rcc_CountryCapital WHERE Cca2 = " & SQ(sCca2), dbOpenSnapshot)
            WriteBeginArray
            Dim bFirstCap As Boolean
            bFirstCap = True
            Do While Not rsCapitals.EOF
                If Not bFirstCap Then WriteComma
                bFirstCap = False
                WriteStringValue Nz(rsCapitals.Fields("Capital").Value, "")
                rsCapitals.MoveNext
            Loop
            WriteEndArray
            rsCapitals.Close

            ' Languages
            WriteComma
            WritePropKey "Languages"
            Set rsLanguages = db.OpenRecordset("SELECT * FROM rcc_CountryLanguage WHERE Cca2 = " & SQ(sCca2) & " ORDER BY Code", dbOpenSnapshot)
            WriteBeginArray
            Dim bFirstLang As Boolean
            bFirstLang = True
            Do While Not rsLanguages.EOF
                If Not bFirstLang Then WriteComma
                bFirstLang = False
                WriteBeginObject
                    WritePropValueString "Code", Nz(rsLanguages.Fields("Code").Value, "")
                    WriteComma
                    WritePropValueString "Name", Nz(rsLanguages.Fields("Name").Value, "")
                WriteEndObject
                rsLanguages.MoveNext
            Loop
            WriteEndArray
            rsLanguages.Close

            ' Currencies
            WriteComma
            WritePropKey "Currencies"
            Set rsCurrencies = db.OpenRecordset("SELECT * FROM rcc_CountryCurrency WHERE Cca2 = " & SQ(sCca2) & " ORDER BY Code", dbOpenSnapshot)
            WriteBeginArray
            Dim bFirstCur As Boolean
            bFirstCur = True
            Do While Not rsCurrencies.EOF
                If Not bFirstCur Then WriteComma
                bFirstCur = False
                WriteBeginObject
                    WritePropValueString "Code", Nz(rsCurrencies.Fields("Code").Value, "")
                    WriteComma
                    WritePropValueString "Name", Nz(rsCurrencies.Fields("Name").Value, "")
                    WriteComma
                    WritePropValueString "Symbol", Nz(rsCurrencies.Fields("Symbol").Value, "")
                WriteEndObject
                rsCurrencies.MoveNext
            Loop
            WriteEndArray
            rsCurrencies.Close

        WriteEndObject
        rsCountry.MoveNext
    Loop
    WriteEndArray

    rsCountry.Close
    Set db = Nothing
End Sub

' ============================================================
' CNS Countries
' ============================================================

Private Sub WriteCnsCountries()
    Dim db As DAO.Database
    Dim rsCountry As DAO.Recordset
    Dim rsCapital As DAO.Recordset
    Dim rsFlag As DAO.Recordset
    Dim rsCities As DAO.Recordset
    Dim sIso2 As String
    Dim bFirstCountry As Boolean

    Set db = CurrentDb()
    Set rsCountry = db.OpenRecordset("SELECT * FROM cns_Country ORDER BY Iso2", dbOpenSnapshot)

    WriteBeginArray
    bFirstCountry = True
    Do While Not rsCountry.EOF
        If Not bFirstCountry Then WriteComma
        bFirstCountry = False

        sIso2 = "" & rsCountry.Fields("Iso2").Value

        WriteBeginObject
            WritePropValueString "Iso2", sIso2
            WriteComma
            WritePropValueString "Country", Nz(rsCountry.Fields("Country").Value, "")
            WriteComma
            WritePropValueString "Iso3", Nz(rsCountry.Fields("Iso3").Value, "")

            ' Capital
            Set rsCapital = db.OpenRecordset("SELECT * FROM cns_CountryCapital WHERE Iso2 = " & SQ(sIso2), dbOpenSnapshot)
            If Not rsCapital.EOF Then
                WriteComma
                WritePropKey "Capital"
                WriteBeginObject
                    WritePropValueString "Name", Nz(rsCapital.Fields("Name").Value, "")
                    WriteComma
                    WritePropValueString "Capital", Nz(rsCapital.Fields("Capital").Value, "")
                WriteEndObject
            End If
            rsCapital.Close

            ' Flag
            Set rsFlag = db.OpenRecordset("SELECT * FROM cns_CountryFlag WHERE Iso2 = " & SQ(sIso2), dbOpenSnapshot)
            If Not rsFlag.EOF Then
                WriteComma
                WritePropKey "Flag"
                WriteBeginObject
                    WritePropValueString "Name", Nz(rsFlag.Fields("Name").Value, "")
                    WriteComma
                    WritePropValueString "Flag", Nz(rsFlag.Fields("Flag").Value, "")
                    WriteComma
                    WritePropValueString "DialCode", Nz(rsFlag.Fields("DialCode").Value, "")
                WriteEndObject
            End If
            rsFlag.Close

            ' Cities (first 20 to avoid enormous output)
            WriteComma
            WritePropKey "Cities"
            Set rsCities = db.OpenRecordset("SELECT TOP 20 * FROM cns_City WHERE Iso2 = " & SQ(sIso2) & " ORDER BY City", dbOpenSnapshot)
            WriteBeginArray
            Dim bFirstCity As Boolean
            bFirstCity = True
            Do While Not rsCities.EOF
                If Not bFirstCity Then WriteComma
                bFirstCity = False
                WriteStringValue Nz(rsCities.Fields("City").Value, "")
                rsCities.MoveNext
            Loop
            WriteEndArray
            rsCities.Close

        WriteEndObject
        rsCountry.MoveNext
    Loop
    WriteEndArray

    rsCountry.Close
    Set db = Nothing
End Sub

' ============================================================
' JSON Helpers
' ============================================================

Private Sub WriteBeginObject()
    m_sJson = m_sJson & "{" & vbCrLf
    m_iIndent = m_iIndent + 1
End Sub

Private Sub WriteEndObject()
    m_iIndent = m_iIndent - 1
    m_sJson = m_sJson & vbCrLf & INDENT() & "}"
End Sub

Private Sub WriteBeginArray()
    m_sJson = m_sJson & "[" & vbCrLf
    m_iIndent = m_iIndent + 1
End Sub

Private Sub WriteEndArray()
    m_iIndent = m_iIndent - 1
    m_sJson = m_sJson & vbCrLf & INDENT() & "]"
End Sub

Private Sub WritePropKey(sKey As String)
    m_sJson = m_sJson & INDENT() & Chr(34) & json_Encode(sKey) & Chr(34) & ": "
End Sub

Private Sub WritePropValueString(sKey As String, sValue As String)
    m_sJson = m_sJson & INDENT() & Chr(34) & json_Encode(sKey) & Chr(34) & ": " & Chr(34) & json_Encode(sValue) & Chr(34)
End Sub

Private Sub WritePropValueLong(sKey As String, lValue As Long)
    m_sJson = m_sJson & INDENT() & Chr(34) & json_Encode(sKey) & Chr(34) & ": " & CStr(lValue)
End Sub

Private Sub WriteStringValue(sValue As String)
    m_sJson = m_sJson & INDENT() & Chr(34) & json_Encode(sValue) & Chr(34)
End Sub

Private Sub WriteComma()
    m_sJson = m_sJson & "," & vbCrLf
End Sub

Private Function INDENT() As String
    INDENT = String(m_iIndent * 2, " ")
End Function

Private Function json_Encode(sValue As String) As String
    Dim sEncoded As String
    sEncoded = sValue
    sEncoded = Replace(sEncoded, "\", "\\")
    sEncoded = Replace(sEncoded, Chr(34), "\" & Chr(34))
    sEncoded = Replace(sEncoded, vbCrLf, "\n")
    sEncoded = Replace(sEncoded, vbCr, "\n")
    sEncoded = Replace(sEncoded, vbLf, "\n")
    sEncoded = Replace(sEncoded, vbTab, "\t")
    json_Encode = sEncoded
End Function

' Helper from SYS__Strings (duplicated here for standalone use)
Private Function SQ(s As String) As String
    SQ = "'" & Replace(s, "'", "''") & "'"
End Function
