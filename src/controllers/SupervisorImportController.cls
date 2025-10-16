VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SupervisorImportController"
Option Explicit

' ===========================================
' CONTROLLER: SupervisorImportController
' ===========================================
' Orchestrates supervisor data import from line workbooks
' into the master yield/efficiency workbook
' ===========================================

Private configHelper As ConfigHelper
Private logger As LogHelper
Private xlHelper As ExcelHelper
Private fileHelper As FileHelper

Public Sub RunImport()
    On Error GoTo ERR_HANDLER

    ' === Initialise helpers ===
    Set configHelper = New ConfigHelper
    Set logger = New LogHelper
    Set xlHelper = New ExcelHelper
    Set fileHelper = New FileHelper

    ' === Start session ===
    logger.StartSession configHelper.GetPath("log_folder")
    logger.Info "=== Supervisor Import Started ==="
    logger.Info "Using static configuration values (no JSON file)."

    ' === Load config ===
    Dim year As Long: year = configHelper.GetYear
    Dim month As String: month = configHelper.GetMonth
    Dim targetDate As Date: targetDate = configHelper.GetTargetDate
    Dim masterPath As String: masterPath = configHelper.GetPath("master_file")
    Dim rootPath As String: rootPath = configHelper.GetPath("supervisor_root")
    Dim shifts As Variant: shifts = configHelper.GetShifts
    Dim lines As Variant: lines = configHelper.GetLines
    Dim colMap As Variant: colMap = configHelper.GetColumns

    Dim shift As Variant, lineName As Variant
    Dim folder As String, fileName As String

    ' === Main loop ===
    For Each shift In shifts
        logger.Info "Processing shift: " & shift
        For Each lineName In lines
            folder = rootPath & shift & "\" & lineName & "\" & month & "\"
            logger.Info "     Folder: " & folder

            fileName = fileHelper.GetLatestFile(folder, targetDate, logger)
            If fileName <> "" Then
                xlHelper.ReadSupervisorData folder & fileName, lineName, shift, masterPath, colMap, logger
            Else
                logger.Warn "     ⚠ No valid Excel file found for " & lineName
            End If
        Next lineName
    Next shift

    ' === Wrap up ===
    logger.Info "✅ Import complete."
    logger.Save
    MsgBox "Supervisor import complete." & vbCrLf & _
           "Log saved to: " & vbCrLf & logger.SessionPath, vbInformation, "Supervisor Import"
    Exit Sub

ERR_HANDLER:
    If Not logger Is Nothing Then
        logger.[Error] "Unhandled error: " & Err.Description
        logger.Save
    Else
        MsgBox "Critical error before logger initialised: " & Err.Description, vbCritical
    End If
End Sub
