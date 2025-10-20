VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MainController"
Option Explicit

Public Sub runYEImport()
    On Error GoTo ERR_HANDLER

    ' Create shared helper instances so controllers share resources 
    Dim cfg As New ConfigHelper
    Dim lg  As New LogHelper
    Dim xl  As New ExcelHelper
    Dim curingXl As New CuringHelper
    Dim fh As New FileHelper
 
    Dim supervisorController As New SupervisorImportController
    'Dim curingController As New CuringImportController

    supervisorController.Init cfg, lg, xl, fh
    'curingController.Init cfg, lg, curingXl, fh

    supervisorController.RunImport
    'curingController.RunImport

    Exit Sub
ERR_HANDLER:
    If Not lg Is Nothing Then
        lg.Error "MainController.runYEImport error " & Err.Number & ": " & Err.Description
    End If
    lg.Info "Execution halted due to error. Please check the log file."
    Exit Sub
End Sub