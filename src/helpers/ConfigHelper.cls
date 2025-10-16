VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigHelper"
Option Explicit

' ===========================================
' HELPER: ConfigHelper
' Centralised configuration (no JSON required)
' ===========================================

' Paths
Private Const BASE_PATH As String = "D:\project\OakPark-Data-YE\"
Private Const MASTER_FILE As String = BASE_PATH & "Master-Data\Master-YE.xlsx"
Private Const SUPERVISOR_ROOT As String = BASE_PATH & "supervisor\"
Private Const LOG_FOLDER As String = BASE_PATH & "Logs\"

' Other configuration
Private Const YEAR_VALUE As Long = 2025
Private Const MONTH_VALUE As String = "May"
Private Const TARGET_DATE As String = "02.05.2025" ' dd.mm.yyyy

Private Shifts As Variant
Private Lines As Variant
Private Columns As Variant

Private Sub Class_Initialize()
    Shifts = Array("day shift", "night shift")
    Lines = Array("Line 1", "Line 2-3", "Line 4")

    ' Column mappings: {SupervisorCol, MasterCol}
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
    
End Sub

Public Function GetYear() As Long: GetYear = YEAR_VALUE: End Function
Public Function GetMonth() As String: GetMonth = MONTH_VALUE: End Function

Public Function GetTargetDate() As Date
    Dim parts() As String
    parts = Split(TARGET_DATE, ".")
    If UBound(parts) = 2 Then
        GetTargetDate = DateSerial(CLng(parts(2)), CLng(parts(1)), CLng(parts(0)))
    Else
        GetTargetDate = Date
    End If
End Function


Public Function GetPath(ByVal key As String) As String
    Select Case LCase(key)
        Case "master_file": GetPath = MASTER_FILE
        Case "supervisor_root": GetPath = SUPERVISOR_ROOT
        Case "log_folder": GetPath = LOG_FOLDER
        Case Else: GetPath = ""
    End Select
End Function

Public Function GetShifts() As Variant: GetShifts = Shifts: End Function
Public Function GetLines() As Variant: GetLines = Lines: End Function
Public Function GetColumns() As Variant: GetColumns = Columns: End Function
