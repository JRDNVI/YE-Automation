Option Explicit

Public Sub ImportSupervisorSheets()
    Dim SupervisorController As New SupervisorImportController
   ' Dim CuringController As New CuringImportController

    SupervisorController.RunImport
   ' CuringController.RunImport
End Sub
