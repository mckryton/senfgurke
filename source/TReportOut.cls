VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TReportOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is the interface for any ReportOut class. A ReportOut class defines the output
'channel for reporting (e.g. the immediate window).

Option Explicit


Public ConsoleOutput As String
Public Sub PrintOut(report_output As String, Optional consecutive As Boolean)
End Sub
