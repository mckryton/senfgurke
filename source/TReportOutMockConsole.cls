VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TReportOutMockConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class sends any report to a string property console_output.
'This way output form the actual test run doesn't get mixed with
'output caused by the step implemations.
'Use TReport.enable_silent_mode to activate this output channel.

Option Explicit

Implements TReportOut

Private console_output As String

Public Property Get TReportOut_ConsoleOutput() As String
    TReportOut_ConsoleOutput = console_output
End Property

Public Property Let TReportOut_ConsoleOutput(ByVal updated_console_output As String)
    console_output = updated_console_output
End Property

Public Sub TReportOut_PrintOut(report_output As String, Optional consecutive As Boolean)
     
    If IsMissing(consecutive) Then consecutive = False
    console_output = console_output & report_output
    If Not consecutive Then console_output = console_output & vbLf
End Sub

Private Sub Class_Initialize()
    console_output = vbNullString
End Sub
