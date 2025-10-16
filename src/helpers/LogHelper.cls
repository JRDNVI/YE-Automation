VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogHelper"
Option Explicit

' ===========================================
' HELPER: LogHelper
' ===========================================

Private logFile As Integer
Private buffer As String
Private logFolder As String
Public SessionPath As String

Public Sub StartSession(ByVal folder As String)
    On Error Resume Next
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    On Error GoTo 0

    logFolder = folder
    SessionPath = folder & "YE_Log_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    logFile = FreeFile
    Open SessionPath For Output As #logFile
    Info "Log session started at " & Now
End Sub

Public Sub Info(ByVal msg As String): WriteLine "INFO", msg: End Sub
Public Sub Warn(ByVal msg As String): WriteLine "WARN", msg: End Sub
Public Sub Error(ByVal msg As String): WriteLine "ERROR", msg: End Sub

Private Sub WriteLine(ByVal level As String, ByVal msg As String)
    Dim line As String
    line = Format(Now, "hh:nn:ss") & " [" & level & "] " & msg
    Debug.Print line
    buffer = buffer & line & vbCrLf
    Application.StatusBar = msg
    DoEvents
    If logFile <> 0 Then Print #logFile, line
End Sub

Public Sub Save()
    If logFile <> 0 Then
        Info "Log session closed."
        Close #logFile
        logFile = 0
    End If
End Sub
