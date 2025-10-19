VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileHelper"
Option Explicit

' ===========================================
' HELPER: FileHelper
' ===========================================

Public Function GetLatestFile(ByVal folder As String, _
                              ByVal targetDate As Date, _
                              ByRef logger As LogHelper) As String
    On Error GoTo EH

    Dim targetFile As String, fileName As String
    Dim latestFile As String, latestDate As Date, thisDate As Date
    Dim parts() As String, base As String

    If Right(folder, 1) <> "\" Then folder = folder & "\"

    targetFile = Format(targetDate, "dd.mm.yyyy") & ".xlsx"
    If Len(Dir(folder & targetFile)) > 0 Then
        logger.Info "Exact match found: " & folder & targetFile
        GetLatestFile = targetFile
        Exit Function
    End If

    If latestFile <> "" Then
        logger.Info "Using latest available file: " & folder & latestFile
    End If

    GetLatestFile = latestFile
    Exit Function

EH:
    logger.[Error] "FileHelper.GetLatestFile: " & Err.Description
    GetLatestFile = ""
End Function
