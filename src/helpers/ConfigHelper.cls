VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigHelper"
Option Explicit

' ===========================================
' HELPER: ConfigHelper
' Centralised configuration For Y&E Automation
' Includes file paths, date settings, column mappings, etc.
' If any changes are needed, they should be made here
' ===========================================

Private Const MASTER_FILE As String = "\\Oak-w2k22-dc01\Department\Production\Production\Yields\2025\Yields and Efficiencies 2025 - Cahir.xlsb"
private Const MASTER_DOWNTIME_FILE As String = "\\Oak-w2k22-dc01\Department\Production\Production\Yields\2025\Downtime Master File 2025 (New).xlsx"
Private Const SUPERVISOR_ROOT As String = "\\Oak-w2k22-dc01\Department\Supervisor\"
Private Const CURING_ROOT As String = SUPERVISOR_ROOT & "Curing Hall\"
Private Const LOG_FOLDER As String = "C:\Users\coadyj\projects\YE-Automation\Logs\"

' This is the year and month for the import process
Private Const YEAR_VALUE As Long = 2025 ' Year used in folder paths
Private Const MONTH_VALUE As String = "October" ' Month name as used in folder paths
Private Const TARGET_DATE As String = "17.10.2025"  ' File to target (dd.mm.yyyy)

Private Shifts As Variant
Private Lines As Variant
private CuringPrefixes As Variant

Private Columns As Variant
Private Line4_Columns As Variant
private production_downtime_columns As Variant
private maintenance_downtime_columns As Variant
private production_downtime_columns_4 As Variant

Private ThermoColumns As Variant
private SmokingColumns As Variant

Private Sub Class_Initialize()
    Shifts = Array("Day Shift Yield & Efficiency", "Night Shift Yield & Efficiency")
    Lines = Array("Line 1", "Line 2 & 3", "Line 4")
    CuringPrefixes = Array("Smoking & Heat Treating Day By Hour Sheet", "Thermo Day By Hour Sheet.")

    ' Column mappings: { SupervisorCol, MasterCol }
    ' This is for Lines 1, 2-3
    Columns = Array( _
        Array("B", "C"), _
        Array("I", "N"), _
        Array("L", "Q"), _
        Array("R", "M"), _
        Array("S", "I"), _
        Array("T", "L"), _
        Array("U", "J"), _
        Array("X", "U"), _
        Array("Z", "V"), _
        Array("AB", "T") _
    )

    ' Column mappings for Line 4 (different layout)
    Line4_Columns = Array( _
        Array("B", "C"), _
        Array("Q", "T"), _
        Array("I", "N"), _
        Array("L", "Q"), _
        Array("R", "M"), _
        Array("S", "I"), _
        Array("X", "Y"), _
        Array("Z", "AA") _
    )

    production_downtime_columns = Array( _
        Array("B", "D"), _
        Array("C", "E"), _
        Array("H", "F"), _
        Array("K", "G"), _
        Array("M", "I"), _
        Array("R", "K"), _
        Array("T", "L"), _
        Array("V", "M"), _
        Array("I", "G"), _
        Array("X", "N"), _
        Array("Z", "P") _
    )

     production_downtime_columns_4 = Array( _
        Array("B", "B"), _
        Array("C", "C"), _
        Array("H", "F"), _
        Array("I", "G"), _
        Array("K", "H"), _
        Array("M", "I"), _
        Array("M", "J"), _
        Array("R", "K"), _
        Array("V", "M"), _
        Array("X", "O"), _
        Array("Z", "Q") _
    )

    maintenance_downtime_columns = Array( _
        Array("C", "C"), _
        Array("I", "P"), _
        Array("M", "N"), _
        Array("M", "F"), _
        Array("O", "Q") _
    )

    ThermoColumns = Array( _ 
        Array("C", "H"), _
        Array("G", "G"), _
        Array("H", "D") _
    )

    SmokingColumns = Array( _
        Array("B", "G"), _
        Array("C", "F"), _
        Array("E", "E"), _
        Array("A", "C"), _
        Array("I", "H"), _
        Array("J", "I") _
    )
    
End Sub

Public Function GetYear() As Long: GetYear = YEAR_VALUE: End Function
Public Function GetMonth() As String: GetMonth = MONTH_VALUE: End Function

Public Function GetTargetDate() As Date
    Dim parts() As String
    parts = Split(TARGET_DATE, ".")
    If UBound(parts) = 2 Then
        GetTargetDate = DateSerial(CLng(parts(2)), CLng(parts(1)), CLng(parts(0)))
    End If
End Function

Public Function GetPath(ByVal key As String) As String
    Select Case LCase(key)
        Case "master_file": GetPath = MASTER_FILE
        Case "supervisor_root": GetPath = SUPERVISOR_ROOT
        Case "master_downtime_file": GetPath = MASTER_DOWNTIME_FILE
        Case "curing_root": GetPath = CURING_ROOT
        Case "log_folder": GetPath = LOG_FOLDER
        Case Else: GetPath = ""
    End Select
End Function

Public Function GetShifts() As Variant: GetShifts = Shifts: End Function
Public Function GetLines() As Variant: GetLines = Lines: End Function
public Function GetCuringPrefixes() As Variant: GetCuringPrefixes = CuringPrefixes: End Function

Public Function GetColumns() As Variant: GetColumns = Columns: End Function
Public Function GetLine4_Columns() As Variant: GetLine4_Columns = Line4_Columns: End Function
Public Function GetThermoColumns() As Variant: GetThermoColumns = ThermoColumns: End Function
Public Function GetSmokingColumns() As Variant: GetSmokingColumns = SmokingColumns: End Function

Public Function GetProductionDowntimeColumns() As Variant: GetProductionDowntimeColumns = production_downtime_columns: End Function
Public Function GetMaintenanceDowntimeColumns() As Variant: GetMaintenanceDowntimeColumns = maintenance_downtime_columns: End Function
public Function GetProductionDowntimeColumns4() As Variant: GetProductionDowntimeColumns4 = production_downtime_columns_4: End Function
