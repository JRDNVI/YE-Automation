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

    logger.Info "Searching for latest available file..."
    fileName = Dir(folder & "*.xls*")

    Do While fileName <> ""
        If Left(fileName, 2) <> "~$" Then
            base = Left(fileName, InStrRev(fileName, ".") - 1)
            parts = Split(base, ".")
            If UBound(parts) = 2 Then
                On Error Resume Next
                thisDate = DateSerial(CLng(parts(2)), CLng(parts(1)), CLng(parts(0)))
                On Error GoTo 0
                If IsDate(thisDate) Then
                    If thisDate > latestDate Then
                        latestDate = thisDate
                        latestFile = fileName
                    End If
                End If
            End If
        End If
        fileName = Dir
    Loop

    If latestFile <> "" Then
        logger.Info "Using latest available file: " & folder & latestFile
    Else
        logger.Warn "No valid date-formatted files found."
    End If

    GetLatestFile = latestFile
    Exit Function

EH:
    logger.[Error] "FileHelper.GetLatestFile: " & Err.Description
    GetLatestFile = ""
End Function
