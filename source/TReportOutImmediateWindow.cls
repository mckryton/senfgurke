VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TReportOutImmediateWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this class sends any report output to the direct window

Option Explicit

Implements TReportOut

Private console_output As String

Public Property Get TReportOut_ConsoleOutput() As String
    TReportOut_ConsoleOutput = console_output
End Property

Public Property Let TReportOut_ConsoleOutput(ByVal dummy As String)
    'TReportOutImmediateWindow only prints directly to the immediate window
    ' and doesn't persist console output
End Property

Public Sub TReportOut_PrintOut(report_output As String, Optional consecutive As Boolean)
    If IsMissing(consecutive) Then consecutive = False
    If consecutive Then
        Debug.Print report_output;
    Else
        Debug.Print report_output
    End If
End Sub

Private Sub Class_Initialize()
    console_output = "error - TReportOutDirectWindow: can't read from immediate window"
End Sub

