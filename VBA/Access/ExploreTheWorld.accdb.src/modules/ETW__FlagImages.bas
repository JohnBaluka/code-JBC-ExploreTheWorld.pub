Attribute VB_Name = "ETW__FlagImages"
Option Compare Database
Option Explicit

' ============================================================
' ETW__FlagImages
' Local cache for country flag images used by the Word/Excel/
' PowerPoint VBA exports. Mirrors the .NET FlagImageManager:
'   1. Look up the cached PNG under
'      %LOCALAPPDATA%\JBC\ExploreTheWorld\FlagImages\{ISO2}.png
'      (same folder as the .NET file cache, so downloads are shared).
'   2. On a cache miss, derive the Wikimedia PNG thumbnail URL from
'      the SVG flag URL stored in cns_CountryFlag.Flag and download it.
' ============================================================

' Wikimedia only serves a fixed list of thumbnail widths (https://w.wiki/GHai);
' 330 is on the list — arbitrary values such as 320 are rejected with HTTP 400.
Private Const FLAG_THUMB_WIDTH_PX As Long = 330

#If VBA7 Then
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As LongPtr) As Long
#Else
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

' Returns the local PNG file path for the country flag, downloading and caching it
' when necessary. Returns "" when no image is available (no URL, offline, ...).
Public Function Get_FlagImageFilePath(oLog As TextBox, sIso2 As String, sFlagUrl As String) As String
    On Error GoTo ErrHandler
    Get_FlagImageFilePath = ""

    Dim sKey As String
    sKey = UCase$(Trim$(sIso2))
    If Len(sKey) = 0 Then Exit Function

    Dim sPath As String
    sPath = Get_CacheFolder() & sKey & ".png"
    If Len(Dir$(sPath)) > 0 Then
        Get_FlagImageFilePath = sPath
        Exit Function
    End If

    Dim sPngUrl As String
    sPngUrl = Get_PngThumbnailUrl(sFlagUrl)
    If Len(sPngUrl) = 0 Then Exit Function

    If URLDownloadToFile(0, sPngUrl, sPath, 0, 0) = 0 And Len(Dir$(sPath)) > 0 Then
        Get_FlagImageFilePath = sPath
    Else
        APP.AppendLog oLog, "  Flag download failed for " & sKey
    End If
    Exit Function

ErrHandler:
    APP.AppendLog oLog, "  Flag image error for " & sIso2 & ": " & Err.Description
    Get_FlagImageFilePath = ""
End Function

' Converts a Wikimedia SVG URL to its rasterized PNG thumbnail URL, e.g.
'   https://upload.wikimedia.org/wikipedia/commons/d/d4/Flag_of_Israel.svg
'   -> https://upload.wikimedia.org/wikipedia/commons/thumb/d/d4/Flag_of_Israel.svg/330px-Flag_of_Israel.svg.png
' The CountriesNow data mixes Commons URLs (/wikipedia/commons/...) with English
' Wikipedia media URLs (/wikipedia/en/...); both use the same thumbnail scheme.
' Returns the URL unchanged for PNG/JPG URLs, "" when the URL is not usable.
' Mirrors JBC.ExploreTheWorld.CL.FlagImageUrl_Helper.GetPngThumbnailUrl.
Public Function Get_PngThumbnailUrl(sFlagUrl As String) As String
    Const WIKIPEDIA_SEGMENT As String = "/wikipedia/"

    Get_PngThumbnailUrl = ""
    Dim sUrl As String
    sUrl = Trim$(sFlagUrl)
    If Len(sUrl) = 0 Then Exit Function

    Dim sLower As String
    sLower = LCase$(sUrl)
    If Right$(sLower, 4) = ".png" Or Right$(sLower, 4) = ".jpg" Or Right$(sLower, 5) = ".jpeg" Then
        Get_PngThumbnailUrl = sUrl
        Exit Function
    End If
    If Right$(sLower, 4) <> ".svg" Then Exit Function

    ' Split into ".../wikipedia/{project}/" prefix and the media-relative path.
    Dim lSegment As Long
    lSegment = InStr(1, sLower, WIKIPEDIA_SEGMENT, vbTextCompare)
    If lSegment = 0 Then Exit Function

    Dim lProjectStart As Long
    Dim lProjectEnd As Long
    lProjectStart = lSegment + Len(WIKIPEDIA_SEGMENT)
    lProjectEnd = InStr(lProjectStart, sUrl, "/")
    If lProjectEnd = 0 Then Exit Function

    Dim sPrefix As String
    Dim sRelative As String
    Dim sFileName As String
    sPrefix = Left$(sUrl, lProjectEnd)
    sRelative = Mid$(sUrl, lProjectEnd + 1)

    ' Already a thumbnail path - append the sized PNG rendition of the SVG file name.
    If LCase$(Left$(sRelative, 6)) = "thumb/" Then
        sFileName = Get_FileName(sUrl)
        Get_PngThumbnailUrl = sUrl & "/" & FLAG_THUMB_WIDTH_PX & "px-" & sFileName & ".png"
        Exit Function
    End If

    sFileName = Get_FileName(sRelative)
    If Len(sFileName) = 0 Then Exit Function

    Get_PngThumbnailUrl = sPrefix & "thumb/" & sRelative & "/" & FLAG_THUMB_WIDTH_PX & "px-" & sFileName & ".png"
End Function

' Cache folder shared with the .NET hosts; created on first use.
Public Function Get_CacheFolder() As String
    Dim sFolder As String
    sFolder = Environ$("LOCALAPPDATA") & "\JBC\ExploreTheWorld\FlagImages\"
    EnsureFolder Environ$("LOCALAPPDATA") & "\JBC\"
    EnsureFolder Environ$("LOCALAPPDATA") & "\JBC\ExploreTheWorld\"
    EnsureFolder sFolder
    Get_CacheFolder = sFolder
End Function

Private Sub EnsureFolder(sFolder As String)
    If Len(Dir$(sFolder, vbDirectory)) = 0 Then MkDir sFolder
End Sub

Private Function Get_FileName(sUrl As String) As String
    Dim lSlash As Long
    lSlash = InStrRev(sUrl, "/")
    If lSlash > 0 Then
        Get_FileName = Mid$(sUrl, lSlash + 1)
    Else
        Get_FileName = sUrl
    End If
End Function
