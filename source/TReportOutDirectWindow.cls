VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TReportOutDirectWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this class sends any report output to the direct window

Option Explicit

Public console_output As String

Public Sub print_out(report_output As String, Optional consecutive As Boolean)
     
    If IsMissing(consecutive) Then consecutive = False
    If consecutive Then
        Debug.Print report_output;
    Else
        Debug.Print report_output
    End If
End Sub

Private Sub Class_Initialize()
    console_output = "error - TReportOutDirectWindow: can't read direct window"
End Sub
