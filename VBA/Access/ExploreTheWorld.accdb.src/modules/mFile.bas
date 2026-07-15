Attribute VB_Name = "mFile"
Option Explicit

Private m_oFSO As Scripting.FileSystemObject
Private m_oAdoStream As ADODB.Stream

Public Sub WriteAllText(sOutputFilePath As String, sText As String)
    Set m_oFSO = New Scripting.FileSystemObject

    If (m_oFSO.FileExists(sOutputFilePath) = True) Then
        m_oFSO.DeleteFile (sOutputFilePath)
    End If

    Set m_oAdoStream = New ADODB.Stream
    m_oAdoStream.Charset = "UTF-8"
    m_oAdoStream.Open

    m_oAdoStream.WriteText sText

    m_oAdoStream.SaveToFile sOutputFilePath, adSaveCreateOverWrite
    m_oAdoStream.Close

    Debug.Print ("Done")
End Sub

Private Sub WriteLine(sLine As String)
    If sLine = "" Then
        m_oAdoStream.WriteText vbCrLf
    Else
        m_oAdoStream.WriteText sLine, StreamWriteEnum.stWriteLine
    End If
End Sub

Public Sub CopyFileIfNotExists(sSourceFile As String, sTargetFile As String)
    Set m_oFSO = New Scripting.FileSystemObject

    If (m_oFSO.FileExists(sSourceFile) = False) Then
        MsgBox "Source File not found - " & sSourceFile
    End If

    If (m_oFSO.FileExists(sTargetFile) = True) Then
        MsgBox "Target File found, do not overwrite - " & sTargetFile
    End If

    Call m_oFSO.CopyFile(sSourceFile, sTargetFile)
End Sub
