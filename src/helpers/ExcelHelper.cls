VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelHelper"
Option Explicit

' ===========================================
' HELPER: ExcelHelper
' ===========================================

Public Sub ReadSupervisorData(ByVal filePath As String, _
                              ByVal lineName As String, _
                              ByVal shiftName As String, _
                              ByVal masterPath As String, _
                              ByVal colMap As Variant, _
                              ByRef logger As LogHelper)

    Dim wb As Workbook, ws As Worksheet
    Dim i As Long, c As Long
    Dim tempRow() As Variant
    Dim dataList As Collection
    Dim preview As String

    logger.Info "     Opening: " & filePath

    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    On Error Resume Next
    Set wb = Workbooks.Open(filePath, UpdateLinks:=0, ReadOnly:=True)
    On Error GoTo 0
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True

    If wb Is Nothing Then
        logger.[Error] "     ❌ Could not open file."
        Exit Sub
    End If

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Sheets("Yields and Efficiency")
    On Error GoTo 0

    If ws Is Nothing Then
        logger.Warn "     ⚠️ Missing 'Yields and Efficiency' sheet."
        wb.Close False
        Exit Sub
    End If

    ' === Read data ===
    Set dataList = New Collection
    i = 9
    Do While Trim(ws.Range("B" & i).Text) <> ""
        ReDim tempRow(LBound(colMap) To UBound(colMap))
        For c = LBound(colMap) To UBound(colMap)
            tempRow(c) = ws.Range(colMap(c)(0) & i).Value  ' read from supervisor column
        Next c
        dataList.Add tempRow
        i = i + 1
        If i > 2000 Then Exit Do    ' safety
    Loop
    wb.Close False

    logger.Info "     ✅ " & dataList.Count & " rows collected."
    If dataList.Count > 0 Then
        preview = ""
        For c = LBound(dataList(1)) To UBound(dataList(1))
            preview = preview & CStr(dataList(1)(c)) & " | "
        Next c
        logger.Info "          ↳ Sample: " & preview
    End If

    ' === Write to Master ===
    WriteToMaster dataList, lineName, shiftName, colMap, masterPath, logger
End Sub

Private Sub WriteToMaster(ByVal dataList As Collection, _
                          ByVal lineName As String, _
                          ByVal shiftName As String, _
                          ByVal colMap As Variant, _
                          ByVal masterPath As String, _
                          ByRef logger As LogHelper)

    Dim wb As Workbook, ws As Worksheet
    Dim nextRow As Long, dataRow As Variant, c As Long, count As Long

    On Error Resume Next
    Set wb = Workbooks.Open(masterPath)
    On Error GoTo 0

    If wb Is Nothing Then
        logger.[Error] "     ❌ Could not open Master workbook."
        Exit Sub
    End If

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Sheets("Sheet1")
    On Error GoTo 0

    If ws Is Nothing Then
        logger.[Error] "     ⚠️ Missing Sheet1 in Master workbook."
        wb.Close False
        Exit Sub
    End If

    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    For Each dataRow In dataList
        With ws
            .Range("A" & nextRow).Value = Date
            .Range("D" & nextRow).Value = lineName
            .Range("E" & nextRow).Value = shiftName
            For c = LBound(colMap) To UBound(colMap)
                .Range(colMap(c)(1) & nextRow).Value = dataRow(c) ' write to master column
            Next c
        End With
        nextRow = nextRow + 1
        count = count + 1
    Next dataRow

    wb.Close True
    logger.Info "     ✅ " & count & " rows written to Master."
End Sub
