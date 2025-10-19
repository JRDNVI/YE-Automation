VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CuringImportController"
Option Explicit

' ===========================================
' CONTROLLER: SupervisorImportController
' ===========================================
' Orchestrates supervisor data import from line workbooks
' into the master yield/efficiency workbook
' ===========================================

Private configHelper As ConfigHelper
Private logger As LogHelper
Private curingHelper As CuringHelper
Private fileHelper As FileHelper

' Init allows DI from MainController 
Public Sub Init(ByRef cfg As ConfigHelper, ByRef lg As LogHelper, ByRef curingXl As CuringHelper, ByRef fh As FileHelper)
    Set configHelper = cfg
    Set logger = lg
    Set curingHelper = curingXl
    Set fileHelper = fh
End Sub


Public Sub RunImport()
    On Error GoTo ERR_HANDLER

    logger.StartSession configHelper.GetPath("log_folder")
    logger.Info "=== Curing Import Started ==="
    logger.Info "Using static configuration values"

    Dim year As Long: year = configHelper.GetYear
    Dim month As String: month = configHelper.GetMonth
    Dim targetDate As Date: targetDate = configHelper.GetTargetDate

    Dim masterPath As String: masterPath = configHelper.GetPath("master_file")
    Dim rootPath As String: rootPath = configHelper.GetPath("supervisor_root")

    Dim shifts As Variant: shifts = configHelper.GetShifts
    Dim CuringPrefixes As Variant: CuringPrefixes = configHelper.GetCuringPrefixes
    Dim thermoColMap As Variant: thermoColMap = configHelper.GetThermoColumns
    'Dim smokingColMap As Variant: smokingColMap = configHelper.GetSmokingColumns

    Dim shift As Variant, prefix As Variant
    Dim folder As String, fileName As String

    ' Get Columns
    For Each shift In shifts
        For Each prefix In CuringPrefixes
            folder = rootPath & shift & "\" & "Curing Hall" & "\" & year & "\" & month & "\" 
            logger.Info "Folder: " & folder

            ' Find latest file for target date
            fileName = fileHelper.GetLatestCuringFile(folder, targetDate, prefix, logger)
            If fileName <> "" And prefix = "Thermo Day By Hour Sheet." Then
                logger.Info "Found file: " & fileName
                curingHelper.ReadThermoData fileName, masterPath, shift, thermoColMap, targetDate, logger
            Else 
                logger.Info "Found file: " & fileName
                'curingXl.ImportSmokingData folder & fileName, masterPath, logger
            End If
        Next prefix
    Next shift

    Exit Sub
ERR_HANDLER:
    If Not logger Is Nothing Then
        logger.Error "CuringImportController.RunImport error " & Err.Number & ": " & Err.Description
    End If
    Resume Next
End Sub