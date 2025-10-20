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

    Dim fName As String
    Dim latestFile As String
    Dim latestDate As Date
    Dim thisDate As Date
    Dim fullPath As String
    Dim targetFile As String

    ' === Ensure valid folder ===
    If Len(folder) = 0 Then
        logger.Error "GetLatestFile: Folder path is empty."
        Exit Function
    End If

    If Right(folder, 1) <> "\" Then folder = folder & "\"
    If Dir(folder, vbDirectory) = "" Then
        logger.Warn "GetLatestFile: Folder does not exist - " & folder
        Exit Function
    End If

    ' === Look for exact match first ===
    targetFile = Format(targetDate, "dd.mm.yyyy") & ".xlsx"
    If Len(Dir(folder & targetFile)) > 0 Then
        logger.Info "Exact match found: " & folder & targetFile
        GetLatestFile = folder & targetFile
        Exit Function
    End If

    ' === Scan for most recent file ===
    fName = Dir(folder & "*.xls*")
    Do While fName <> ""
        ' Skip temp/hidden/backup files
        If Left(fName, 2) <> "~$" And InStr(1, fName, "backup", vbTextCompare) = 0 Then
            fullPath = folder & fName
            On Error Resume Next
            thisDate = FileDateTime(fullPath)
            If Err.Number = 0 Then
                If thisDate > latestDate Then
                    latestDate = thisDate
                    latestFile = fullPath
                End If
            Else
                logger.Warn "Skipping invalid file: " & fullPath & " (" & Err.Description & ")"
                Err.Clear
            End If
            On Error GoTo EH
        End If
        fName = Dir
    Loop

    ' === Return latest file if available ===
    If Len(latestFile) > 0 Then
        logger.Info "Using latest available file: " & latestFile
        GetLatestFile = latestFile
    Else
        logger.Warn "No valid Excel files found in: " & folder
        GetLatestFile = ""
    End If

    Exit Function

EH:
    logger.[Error] "FileHelper.GetLatestFile error " & Err.Number & ": " & Err.Description
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





