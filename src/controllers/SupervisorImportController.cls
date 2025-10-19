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

Public Sub Init(ByRef cfg As ConfigHelper, ByRef lg As LogHelper, ByRef xl As ExcelHelper, ByRef fh As FileHelper)
    Set configHelper = cfg
    Set logger = lg
    Set xlHelper = xl
    Set fileHelper = fh
End Sub

Public Sub RunImport()
    On Error GoTo ERR_HANDLER

    ' Start log session
    logger.StartSession configHelper.GetPath("log_folder")
    logger.Info "=== Supervisor Import Started ==="
    logger.Info "Using static configuration values"

    ' Load config 
    Dim year As Long: year = configHelper.GetYear
    Dim month As String: month = configHelper.GetMonth
    Dim targetDate As Date: targetDate = configHelper.GetTargetDate

    Dim masterPath As String: masterPath = configHelper.GetPath("master_file")
    Dim rootPath As String: rootPath = configHelper.GetPath("supervisor_root")
    Dim downtimePath As String: downtimePath = configHelper.GetPath("master_downtime_file")
    
    Dim shifts As Variant: shifts = configHelper.GetShifts
    Dim lines As Variant: lines = configHelper.GetLines

    Dim colMap As Variant: colMap = configHelper.GetColumns
    Dim prodCol As Variant: prodCol = configHelper.GetProductionDowntimeColumns
    Dim maintCol As Variant: maintCol = configHelper.GetMaintenanceDowntimeColumns

    Dim shift As Variant, lineName As Variant
    Dim folder As String, fileName As String

    ' Loop Through shifts and lines and create file path
    For Each shift In shifts
        logger.Info "Processing shift: " & shift
        For Each lineName In lines
            folder = rootPath & shift & "\" & lineName & "\" & year & "\" & month & "\" 
            logger.Info "Folder: " & folder


            ' Find latest file for target date
            fileName = fileHelper.GetLatestFile(folder, targetDate, logger)
            If fileName <> "" Then

                ' If Line 4, use different column mappings
                If lineName = "Line 4" Then
                    colMap = configHelper.GetLine4_Columns
                    prodCol = configHelper.GetProductionDowntimeColumns4
                Else
                    colMap = configHelper.GetColumns
                End If

                ' Call Excel helper to read data from supervisor files
                xlHelper.ReadSupervisorData folder & fileName, lineName, shift, masterPath, colMap, logger, targetDate
                xlHelper.ReadProductionDowntime folder & fileName, lineName, shift, downtimePath, prodCol, logger, targetDate
                xlHelper.ReadBreakdownDowntime folder & fileName, lineName, shift, downtimePath, maintCol, logger, targetDate
            Else
                logger.Warn " No valid Excel file found for " & lineName
            End If
        Next lineName
    Next shift

    logger.Info "Import complete."
    logger.Save
    Exit Sub

ERR_HANDLER:
    If Not logger Is Nothing Then
        logger.[Error] "Unhandled error: " & Err.Description
        logger.Save
    Else
        MsgBox "Critical error before logger initialised: " & Err.Description, vbCritical
    End If
End Sub
