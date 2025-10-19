VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelHelper"
Option Explicit

' ===========================================
' HELPER: ExcelHelper
' This class contains methods to read data from
' supervisor line workbooks and write to master workbooks.
' ===========================================

Public Sub ReadSupervisorData(ByVal filePath As String, _
                              ByVal lineName As String, _
                              ByVal shiftName As String, _
                              ByVal masterPath As String, _
                              ByVal colMap As Variant, _
                              ByRef logger As LogHelper, _
                              ByVal targetDate As Date)

    Dim wb As Workbook, ws As Worksheet
    Dim i As Long, c As Long
    Dim tempRow() As Variant
    Dim dataList As Collection
    Dim preview As String

    logger.Info " Opening: " & filePath

    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    On Error Resume Next
    Set wb = Workbooks.Open(filePath, UpdateLinks:=0, ReadOnly:=True)
    On Error GoTo 0
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True

    If wb Is Nothing Then
        logger.[Error] " Could not open file."
        Exit Sub
    End If

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Sheets("Yields and Efficiency")
    On Error GoTo 0

    If ws Is Nothing Then
        logger.Warn " Missing 'Yields and Efficiency' sheet."
        wb.Close False
        Exit Sub
    End If

    Dim supervisor As String
    supervisor = Trim(ws.Range("J7").Text)
    logger.Info " Supervisor: " & supervisor

    Set dataList = New Collection
    i = 9  ' starting row for data

    Do While Not IsEmpty(ws.Range("I" & i).Value) And ws.Range("I" & i).Value <> 0
        ReDim tempRow(LBound(colMap) To UBound(colMap))
        For c = LBound(colMap) To UBound(colMap)
            tempRow(c) = ws.Range(colMap(c)(0) & i).Value  ' read from supervisor column
        Next c
        dataList.Add tempRow
        i = i + 1
        If i > 2000 Then Exit Do  ' safety stop
    Loop

    wb.Close False


    logger.Info dataList.Count & " rows collected."

    ' === Write to Master ===
    WriteToMaster dataList, lineName, shiftName, colMap, masterPath, logger, targetDate, supervisor
End Sub

Public Sub ReadProductionDowntime(ByVal filePath As String, _
                                  ByVal lineName As String, _
                                  ByVal shiftName As String, _
                                  ByVal downtimePath As String, _
                                  ByVal colMap As Variant, _
                                  ByRef logger As LogHelper, _
                                  ByVal targetDate As Date)

    Dim wb As Workbook, ws As Worksheet
    Dim i As Long, c As Long
    Dim dataList As Collection
    Dim tempRow() As Variant
    Dim foundCount As Long, startRow As Long
    Dim lastRow As Long
    Dim cell As Range

    logger.Info " Opening for Production Downtime: " & filePath

    On Error Resume Next
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)
    On Error GoTo 0

    If wb Is Nothing Then
        logger.Error " Could not open file for downtime extraction."
        Exit Sub
    End If

    On Error Resume Next
    Set ws = wb.Sheets("Yields and Efficiency")
    On Error GoTo 0

    If ws Is Nothing Then
        logger.Warn " Missing 'Yields and Efficiency' sheet."
        wb.Close False
        Exit Sub
    End If

    ' === Find the SECOND occurrence of the word "Product" in column C ===
    foundCount = 0
    For Each cell In ws.Range("C1:C" & ws.Rows.Count)
        If LCase(Trim(cell.Text)) = "product" Then
            foundCount = foundCount + 1
            If foundCount = 2 Then
                startRow = cell.Row + 1
                Exit For
            End If
        End If
    Next cell

    If startRow = 0 Then
        logger.Warn " Could not find second 'Product' header – skipping downtime import."
        wb.Close False
        Exit Sub
    End If

    logger.Info " Found 'Production Downtime' starting at row " & startRow

    ' === Read data from column B onwards until blank, 0, N/A, or Total ===
    Set dataList = New Collection
    i = startRow

    Do While True
        Dim val As Variant
        val = ws.Range("C" & i).Value

        ' === SAFETY: Skip rows with formula errors or invalid values ===
        If IsError(val) Then Exit Do
        If IsEmpty(val) Then Exit Do
        If Trim(val & "") = "" Then Exit Do
        If IsNumeric(val) And val = 0 Then Exit Do
        If LCase(val) = "total" Or LCase(val) = "n/a" Then Exit Do

        ' === Build row ===
        ReDim tempRow(LBound(colMap) To UBound(colMap))
        For c = LBound(colMap) To UBound(colMap)
            tempRow(c) = ws.Range(colMap(c)(0) & i).Value
        Next c
        dataList.Add tempRow

        i = i + 1
        If i > 2000 Then Exit Do
    Loop


    wb.Close False

    logger.Info dataList.Count & " Downtime rows collected."

    'Write to Master workbook 
    WriteProductionDowntimeToMaster dataList, lineName, shiftName, colMap, downtimePath, logger, targetDate
End Sub

Public Sub ReadBreakdownDowntime(ByVal filePath As String, _
                                 ByVal lineName As String, _
                                 ByVal shiftName As String, _
                                 ByVal downtimePath As String, _
                                 ByVal colMap As Variant, _
                                 ByRef logger As LogHelper, _
                                 ByVal targetDate As Date)

    Dim wb As Workbook, ws As Worksheet
    Dim i As Long, c As Long
    Dim dataList As Collection
    Dim tempRow() As Variant
    Dim foundCount As Long, startRow As Long
    Dim val As Variant
    Dim cell As Range

    logger.Info "Opening for Breakdown Downtime: " & filePath

    'Open file safely 
    On Error Resume Next
    Set wb = Workbooks.Open(CStr(filePath), ReadOnly:=True)
    On Error GoTo 0

    If wb Is Nothing Then
        logger.Error "Could not open workbook."
        Exit Sub
    End If

    On Error Resume Next
    Set ws = wb.Sheets("Yields and Efficiency")
    On Error GoTo 0

    If ws Is Nothing Then
        logger.Warn "Missing 'Yields and Efficiency' sheet."
        wb.Close False
        Exit Sub
    End If

    ' Find THIRD "Product" occurrence (start of downtime) 
    foundCount = 0
    Dim lastScanRow As Long
    lastScanRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    For Each cell In ws.Range("C1:C" & lastScanRow)
        If LCase(Trim(CStr(cell.Value))) = "product" Then
            foundCount = foundCount + 1
            If foundCount = 3 Then
                startRow = cell.Row + 1
                Exit For
            End If
        End If
    Next cell

    If startRow = 0 Then
        logger.Warn "Could not find third 'Product' header – skipping downtime."
        wb.Close False
        Exit Sub
    End If

    logger.Info "Found Breakdown Downtime start at row " & startRow

    Set dataList = New Collection
    i = startRow

    Do While True
    
        val = ws.Range("M" & i).Value

        If IsError(val) Then Exit Do
        If IsEmpty(val) Then Exit Do
        If LCase(Trim(CStr(ws.Range("C" & i).Value))) = "total" Then Exit Do
   
        ReDim tempRow(LBound(colMap) To UBound(colMap))
        For c = LBound(colMap) To UBound(colMap)
            tempRow(c) = ws.Range(colMap(c)(0) & i).Value
        Next c
        dataList.Add tempRow

        i = i + 1
        If i > 2000 Then Exit Do
    Loop

    wb.Close False

    logger.Info "Total Breakdown downtime rows collected: " & dataList.Count

   
    WriteBreakdownDowntimeToMaster dataList, lineName, shiftName, colMap, downtimePath, logger, targetDate
End Sub

Private Sub WriteBreakdownDowntimeToMaster(ByVal dataList As Collection, _
                                            ByVal lineName As String, _
                                            ByVal shiftName As String, _
                                            ByVal colMap As Variant, _
                                            ByVal downtimePath As String, _
                                            ByRef logger As LogHelper, _
                                            ByVal targetDate As Date)

    Dim wb As Workbook, ws As Worksheet
    Dim nextRow As Long, dataRow As Variant, c As Long
    Dim count As Long

    logger.Info "Writing Maintenance downtime to master file..."

    On Error Resume Next
    Set wb = Workbooks.Open(downtimePath)
    On Error GoTo 0

    If wb Is Nothing Then
        logger.Error "Could not open Master workbook for downtime write."
        Exit Sub
    End If

    'Locate or create target sheet
    On Error Resume Next
    Set ws = wb.Sheets("Sheet1")
    On Error GoTo 0

    'Find the next available row 
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    For Each dataRow In dataList
        With ws
            .Range("A" & nextRow).Value = targetDate
            .Range("B" & nextRow).Value = lineName
            .Range("C" & nextRow).Value = shiftName
            .Range("O" & nextRow).Value = "Maintenance"  

            For c = LBound(colMap) To UBound(colMap)
                .Range(colMap(c)(1) & nextRow).Value = dataRow(c)
            Next c
        End With

        nextRow = nextRow + 1
        count = count + 1
    Next dataRow

    wb.Close SaveChanges:=True
    logger.Info count & " production downtime rows written to Master file."
End Sub



Private Sub WriteProductionDowntimeToMaster(ByVal dataList As Collection, _
                                            ByVal lineName As String, _
                                            ByVal shiftName As String, _
                                            ByVal colMap As Variant, _
                                            ByVal downtimePath As String, _
                                            ByRef logger As LogHelper, _
                                            ByVal targetDate As Date)

    Dim wb As Workbook, ws As Worksheet
    Dim nextRow As Long, dataRow As Variant, c As Long
    Dim count As Long

    logger.Info "Writing production downtime to master file..."

    ' === Open master workbook ===
    On Error Resume Next
    Set wb = Workbooks.Open(downtimePath)
    On Error GoTo 0

    If wb Is Nothing Then
        logger.Error "Could not open Master workbook for downtime write."
        Exit Sub
    End If

    'Locate or create target sheet
    On Error Resume Next
    Set ws = wb.Sheets("Sheet1")
    On Error GoTo 0

    'Find the next available row 
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1


    For Each dataRow In dataList
        With ws
            .Range("A" & nextRow).Value = targetDate
            .Range("B" & nextRow).Value = lineName
            .Range("C" & nextRow).Value = shiftName
            .Range("O" & nextRow).Value = "Production"  

            For c = LBound(colMap) To UBound(colMap)
                .Range(colMap(c)(1) & nextRow).Value = dataRow(c)
            Next c
        End With

        nextRow = nextRow + 1
        count = count + 1
    Next dataRow

    wb.Close SaveChanges:=True
    logger.Info count & " production downtime rows written to Master file."
End Sub


Private Sub WriteToMaster(ByVal dataList As Collection, _
                          ByVal lineName As String, _
                          ByVal shiftName As String, _
                          ByVal colMap As Variant, _
                          ByVal masterPath As String, _
                          ByRef logger As LogHelper, _ 
                          ByVal targetDate As Date, _
                          ByVal supervisor As String)

    Dim wb As Workbook, ws As Worksheet
    Dim nextRow As Long, dataRow As Variant, c As Long, count As Long

    On Error Resume Next
    Set wb = Workbooks.Open(masterPath)
    On Error GoTo 0

    If wb Is Nothing Then
        logger.[Error] "Could not open Master workbook."
        Exit Sub
    End If

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Sheets("Yield & Efficiencies Lines")
    On Error GoTo 0

    If ws Is Nothing Then
        logger.[Error] "Missing Sheet1 in Master workbook."
        wb.Close False
        Exit Sub
    End If

    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    For Each dataRow In dataList
        With ws
            .Range("A" & nextRow).Value = targetDate
            .Range("D" & nextRow).Value = lineName
            .Range("E" & nextRow).Value = shiftName
            .Range("R" & nextRow).Value = supervisor
            For c = LBound(colMap) To UBound(colMap)
                .Range(colMap(c)(1) & nextRow).Value = dataRow(c) 
            Next c
        End With
        nextRow = nextRow + 1
        count = count + 1
    Next dataRow

    wb.Close True
    logger.Info count & " Rows written to Master."
End Sub
