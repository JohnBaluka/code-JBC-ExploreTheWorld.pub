Attribute VB_Name = "ETW__cns_API"
Option Explicit

' ============================================================
' ETW__cns_API
' Loads and clears all CountriesNow.space (cns_) tables by
' calling the countriesnow.space REST API directly.
'
' API base: https://countriesnow.space/api/v0.1
' Tables managed (cns_ prefix):
'   cns_Country, cns_City,
'   cns_CountryCapital, cns_CountryFlag,
'   cns_CountryPopulation, cns_PopulationCount,
'   cns_CountryStates, cns_CountryState
' ============================================================

Private Const CNS_BASE_URL As String = "https://countriesnow.space/api/v0.1"

' -------------------------------------------------------
' Load
' Clears all cns_ tables then fetches fresh data from the
' countriesnow.space API, populating all cns_ tables.
' Calls Screen.ActiveForm.Requery when finished.
' -------------------------------------------------------
Public Sub Load(oLog As TextBox)
    APP.AppendLog oLog, "=== CNS Load started ==="

    On Error GoTo ErrHandler

    ' --- 1. Clear all tables ---
    ClearTables oLog

    ' --- 2. Countries + Cities ---
    APP.AppendLog oLog, "Fetching countries + cities..."
    Dim sJson As String
    sJson = HttpGet(oLog, CNS_BASE_URL & "/countries")
    If Len(sJson) = 0 Then Err.Raise vbObjectError + 1, , "Empty response from /countries"

    Dim oResponse As Object
    Set oResponse = JsonConverter.ParseJson(sJson)
    If oResponse("error") = True Then Err.Raise vbObjectError + 2, , "API error: " & oResponse("msg")

    Dim oData As Object
    Set oData = oResponse("data")

    Dim db As DAO.Database
    Set db = CurrentDb()

    Dim oCountry As Object
    Dim rsCountry As DAO.Recordset
    Dim rsCity As DAO.Recordset
    Set rsCountry = db.OpenRecordset("cns_Country", dbOpenDynaset)
    Set rsCity = db.OpenRecordset("cns_City", dbOpenDynaset)

    Dim lCountryCount As Long
    Dim lCityCount As Long

    For Each oCountry In oData
        Dim sIso2 As String
        sIso2 = Nz(oCountry("iso2"), "")
        If Len(sIso2) = 0 Then GoTo NextCountry

        rsCountry.AddNew
        rsCountry.Fields("GUID").Value = SYS__Guid.NewAccessGuidString()
        rsCountry.Fields("Iso2").Value = sIso2
        rsCountry.Fields("Country").Value = Nz(oCountry("country"), "")
        rsCountry.Fields("Iso3").Value = Nz(oCountry("iso3"), "")
        rsCountry.Update
        lCountryCount = lCountryCount + 1

        ' Cities (array of strings)
        If Not oCountry("cities") Is Nothing Then
            Dim oCities As Collection
            Set oCities = oCountry("cities")
            Dim vCity As Variant
            For Each vCity In oCities
                rsCity.AddNew
                rsCity.Fields("GUID").Value = SYS__Guid.NewAccessGuidString()
                rsCity.Fields("Iso2").Value = sIso2
                rsCity.Fields("City").Value = "" & vCity
                rsCity.Update
                lCityCount = lCityCount + 1
            Next vCity
        End If

NextCountry:
    Next oCountry

    rsCountry.Close
    rsCity.Close
    APP.AppendLog oLog, "  Countries: " & lCountryCount & ", Cities: " & lCityCount

    ' Build lookup of valid Iso2 values for FK-safe child inserts
    Dim oValidIso2 As Object
    Set oValidIso2 = CreateObject("Scripting.Dictionary")
    oValidIso2.CompareMode = vbTextCompare
    Dim rsIso2Check As DAO.Recordset
    Set rsIso2Check = db.OpenRecordset("SELECT Iso2 FROM cns_Country", dbOpenForwardOnly)
    Do While Not rsIso2Check.EOF
        oValidIso2(rsIso2Check.Fields("Iso2").Value) = True
        rsIso2Check.MoveNext
    Loop
    rsIso2Check.Close

    ' --- 3. Capitals ---
    APP.AppendLog oLog, "Fetching capitals..."
    sJson = HttpGet(oLog, CNS_BASE_URL & "/countries/capital")
    Set oResponse = JsonConverter.ParseJson(sJson)
    Set oData = oResponse("data")

    Dim rsCapital As DAO.Recordset
    Set rsCapital = db.OpenRecordset("cns_CountryCapital", dbOpenDynaset)
    Dim lCapCount As Long
    Dim oSeenCapital As Object
    Set oSeenCapital = CreateObject("Scripting.Dictionary")
    oSeenCapital.CompareMode = vbTextCompare

    Dim oCapItem As Object
    For Each oCapItem In oData
        sIso2 = Nz(oCapItem("iso2"), "")
        If Len(sIso2) = 0 Then GoTo NextCapital
        If Not oValidIso2.Exists(sIso2) Then GoTo NextCapital
        If oSeenCapital.Exists(sIso2) Then GoTo NextCapital
        rsCapital.AddNew
        rsCapital.Fields("GUID").Value = SYS__Guid.NewAccessGuidString()
        rsCapital.Fields("Iso2").Value = sIso2
        rsCapital.Fields("Name").Value = Nz(oCapItem("name"), "")
        rsCapital.Fields("Capital").Value = Nz(oCapItem("capital"), "")
        rsCapital.Update
        oSeenCapital(sIso2) = True
        lCapCount = lCapCount + 1
NextCapital:
    Next oCapItem
    rsCapital.Close
    APP.AppendLog oLog, "  Capitals: " & lCapCount

    ' --- 4. Flags ---
    APP.AppendLog oLog, "Fetching flags..."
    sJson = HttpGet(oLog, CNS_BASE_URL & "/countries/flag/images")
    Set oResponse = JsonConverter.ParseJson(sJson)
    Set oData = oResponse("data")

    Dim rsFlag As DAO.Recordset
    Set rsFlag = db.OpenRecordset("cns_CountryFlag", dbOpenDynaset)
    Dim lFlagCount As Long
    Dim oSeenFlag As Object
    Set oSeenFlag = CreateObject("Scripting.Dictionary")
    oSeenFlag.CompareMode = vbTextCompare

    Dim oFlagItem As Object
    For Each oFlagItem In oData
        sIso2 = Nz(oFlagItem("iso2"), "")
        If Len(sIso2) = 0 Then GoTo NextFlag
        If Not oValidIso2.Exists(sIso2) Then GoTo NextFlag
        If oSeenFlag.Exists(sIso2) Then GoTo NextFlag
        rsFlag.AddNew
        rsFlag.Fields("GUID").Value = SYS__Guid.NewAccessGuidString()
        rsFlag.Fields("Iso2").Value = sIso2
        rsFlag.Fields("Name").Value = Nz(oFlagItem("name"), "")
        rsFlag.Fields("Flag").Value = Nz(oFlagItem("flag"), "")
        rsFlag.Fields("DialCode").Value = Nz(oFlagItem("dial_code"), "")
        rsFlag.Update
        oSeenFlag(sIso2) = True
        lFlagCount = lFlagCount + 1
NextFlag:
    Next oFlagItem
    rsFlag.Close
    APP.AppendLog oLog, "  Flags: " & lFlagCount

    ' --- 5. Population ---
    APP.AppendLog oLog, "Fetching population..."
    sJson = HttpGet(oLog, CNS_BASE_URL & "/countries/population")
    Set oResponse = JsonConverter.ParseJson(sJson)
    Set oData = oResponse("data")

    Dim rsPopulation As DAO.Recordset
    Dim rsPopCount As DAO.Recordset
    Set rsPopulation = db.OpenRecordset("cns_CountryPopulation", dbOpenDynaset)
    Set rsPopCount = db.OpenRecordset("cns_PopulationCount", dbOpenDynaset)
    Dim lPopCountries As Long
    Dim lPopCounts As Long

    Dim oPopItem As Object
    For Each oPopItem In oData
        Dim sPopGuid As String
        sPopGuid = SYS__Guid.NewAccessGuidString()
        rsPopulation.AddNew
        rsPopulation.Fields("GUID").Value = sPopGuid
        rsPopulation.Fields("Country").Value = Nz(oPopItem("country"), "")
        rsPopulation.Fields("Code").Value = Nz(oPopItem("code"), "")
        rsPopulation.Fields("Iso3").Value = Nz(oPopItem("iso3"), "")
        rsPopulation.Update
        lPopCountries = lPopCountries + 1

        If Not oPopItem("populationCounts") Is Nothing Then
            Dim oPopCounts As Collection
            Set oPopCounts = oPopItem("populationCounts")
            Dim oPopCount As Object
            For Each oPopCount In oPopCounts
                rsPopCount.AddNew
                rsPopCount.Fields("GUID").Value = SYS__Guid.NewAccessGuidString()
                rsPopCount.Fields("CountryPopulation_GUID").Value = sPopGuid
                rsPopCount.Fields("Year").Value = CLng(Nz(oPopCount("year"), 0))
                rsPopCount.Fields("Value").Value = CDbl(Nz(oPopCount("value"), 0))
                rsPopCount.Update
                lPopCounts = lPopCounts + 1
            Next oPopCount
        End If
    Next oPopItem
    rsPopulation.Close
    rsPopCount.Close
    APP.AppendLog oLog, "  Population entries: " & lPopCountries & ", counts: " & lPopCounts

    ' --- 6. States ---
    APP.AppendLog oLog, "Fetching states..."
    sJson = HttpGet(oLog, CNS_BASE_URL & "/countries/states")
    Set oResponse = JsonConverter.ParseJson(sJson)
    Set oData = oResponse("data")

    Dim rsStates As DAO.Recordset
    Dim rsState As DAO.Recordset
    Set rsStates = db.OpenRecordset("cns_CountryStates", dbOpenDynaset)
    Set rsState = db.OpenRecordset("cns_CountryState", dbOpenDynaset)
    Dim lStatesCount As Long
    Dim lStateCount As Long

    Dim oStatesItem As Object
    For Each oStatesItem In oData
        Dim sStatesGuid As String
        sStatesGuid = SYS__Guid.NewAccessGuidString()
        rsStates.AddNew
        rsStates.Fields("GUID").Value = sStatesGuid
        rsStates.Fields("Name").Value = Nz(oStatesItem("name"), "")
        rsStates.Fields("Iso3").Value = Nz(oStatesItem("iso3"), "")
        rsStates.Update
        lStatesCount = lStatesCount + 1

        If Not oStatesItem("states") Is Nothing Then
            Dim oStates As Collection
            Set oStates = oStatesItem("states")
            Dim oStateItem As Object
            For Each oStateItem In oStates
                rsState.AddNew
                rsState.Fields("GUID").Value = SYS__Guid.NewAccessGuidString()
                rsState.Fields("CountryStates_GUID").Value = sStatesGuid
                rsState.Fields("Name").Value = Nz(oStateItem("name"), "")
                rsState.Fields("StateCode").Value = Nz(oStateItem("state_code"), "")
                rsState.Update
                lStateCount = lStateCount + 1
            Next oStateItem
        End If
    Next oStatesItem
    rsStates.Close
    rsState.Close
    APP.AppendLog oLog, "  State groups: " & lStatesCount & ", states: " & lStateCount

    Set db = Nothing

    APP.AppendLog oLog, "=== CNS Load complete ==="
    Screen.ActiveForm.Requery
    Exit Sub

ErrHandler:
    APP.AppendLog oLog, "ERROR: " & Err.Description
End Sub

' -------------------------------------------------------
' Clear
' Deletes all records from all cns_ tables in FK-safe order.
' Calls Screen.ActiveForm.Requery when finished.
' -------------------------------------------------------
Public Sub Clear(oLog As TextBox)
    APP.AppendLog oLog, "=== CNS Clear started ==="
    ClearTables oLog
    APP.AppendLog oLog, "=== CNS Clear complete ==="
    Screen.ActiveForm.Requery
End Sub

' -------------------------------------------------------
' ClearTables  (private)
' Deletes all rows from cns_ tables in FK-safe order.
' -------------------------------------------------------
Private Sub ClearTables(oLog As TextBox)
    Dim db As DAO.Database
    Set db = CurrentDb()

    APP.AppendLog oLog, "  Deleting cns_PopulationCount..."
    db.Execute "DELETE FROM cns_PopulationCount", dbFailOnError

    APP.AppendLog oLog, "  Deleting cns_CountryPopulation..."
    db.Execute "DELETE FROM cns_CountryPopulation", dbFailOnError

    APP.AppendLog oLog, "  Deleting cns_CountryState..."
    db.Execute "DELETE FROM cns_CountryState", dbFailOnError

    APP.AppendLog oLog, "  Deleting cns_CountryStates..."
    db.Execute "DELETE FROM cns_CountryStates", dbFailOnError

    APP.AppendLog oLog, "  Deleting cns_City..."
    db.Execute "DELETE FROM cns_City", dbFailOnError

    APP.AppendLog oLog, "  Deleting cns_CountryCapital..."
    db.Execute "DELETE FROM cns_CountryCapital", dbFailOnError

    APP.AppendLog oLog, "  Deleting cns_CountryFlag..."
    db.Execute "DELETE FROM cns_CountryFlag", dbFailOnError

    APP.AppendLog oLog, "  Deleting cns_Country..."
    db.Execute "DELETE FROM cns_Country", dbFailOnError

    Set db = Nothing
End Sub

' -------------------------------------------------------
' HttpGet  (private)
' Synchronous HTTP GET. Logs the URL, response status, and
' on non-200 logs up to 500 chars of the response body.
' -------------------------------------------------------
Private Function HttpGet(oLog As TextBox, sUrl As String) As String
    APP.AppendLog oLog, "  GET " & sUrl
    Dim oHttp As MSXML2.XMLHTTP
    Set oHttp = New MSXML2.XMLHTTP
    oHttp.Open "GET", sUrl, False
    oHttp.setRequestHeader "Accept", "application/json"
    oHttp.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    oHttp.send
    APP.AppendLog oLog, "  HTTP " & oHttp.Status & " " & oHttp.statusText
    If oHttp.Status = 200 Then
        HttpGet = oHttp.responseText
    Else
        APP.AppendLog oLog, "  Response: " & Left(oHttp.responseText, 500)
        Err.Raise vbObjectError + 100, , "HTTP " & oHttp.Status & " from " & sUrl
    End If
    Set oHttp = Nothing
End Function
