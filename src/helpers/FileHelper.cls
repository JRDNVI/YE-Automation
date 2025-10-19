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


Public Function GetLatestCuringFile(ByVal folder As String, _
                                    ByVal targetDate As Date, _
                                    ByVal prefix As String, _
                                    ByRef logger As LogHelper) As String
    On Error GoTo EH

    Dim fName As String
    Dim fullPath As String
    Dim found As Boolean
    Dim cleanName As String
    Dim expectedLong As String, expectedShort As String

    ' Ensure trailing backslash
    If Right(folder, 1) <> "\" Then folder = folder & "\"

    ' Build both possible date strings
    expectedLong = Format(targetDate, "dd.mm.yyyy")
    expectedShort = Format(targetDate, "dd.mm.yy")

    ' Normalise prefix
    prefix = UCase(Trim(prefix))

    fName = Dir(folder & "*.xls*")
    Do While fName <> ""
        ' Skip temp/backup files
        If Left(fName, 2) <> "~$" And InStr(1, fName, "backup", vbTextCompare) = 0 Then
            
            ' Clean filename
            cleanName = UCase(Replace(Replace(fName, "_", " "), "  ", " "))
            cleanName = Trim(Replace(cleanName, "  ", " ")) ' collapse double spaces

            ' Check prefix at start, and either date format
            If Left(cleanName, Len(prefix)) = prefix And _
               (InStr(1, cleanName, expectedLong, vbTextCompare) > 0 Or _
                InStr(1, cleanName, expectedShort, vbTextCompare) > 0) Then

                fullPath = folder & fName
                logger.Info "Curing file found for prefix '" & prefix & _
                            "' and date " & expectedLong & ": " & fullPath
                GetLatestCuringFile = fullPath
                found = True
                Exit Do
            End If
        End If
        fName = Dir
    Loop

    If Not found Then
        logger.Warn "No curing file found for prefix '" & prefix & _
                    "' and date " & expectedLong & " in folder: " & folder
        GetLatestCuringFile = ""
    End If

    Exit Function

EH:
    logger.[Error] "FileHelper.GetLatestCuringFile error " & Err.Number & ": " & Err.Description
    GetLatestCuringFile = ""
End Function





