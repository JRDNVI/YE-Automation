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
Private xlHelper As ExcelHelper
Private fileHelper As FileHelper

