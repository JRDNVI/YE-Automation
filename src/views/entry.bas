Option Explicit

Public Sub ImportSupervisorSheets()
    Dim controller As New SupervisorImportController
    controller.RunImport
End Sub
