VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_logger As Logger
Dim m_reporter As TReport

Public Property Get Log() As Logger
    If m_logger Is Nothing Then
        Set m_logger = New Logger
    End If
    Set Log = m_logger
End Property

Public Sub set_report_format(report_format As String)
    Select Case report_format
        Case "verbose", "v"
            Reporter.Report_Formatter = New TReportFormatterVerbose
        Case "progress", "p"
            Reporter.Report_Formatter = New TReportFormatterProgress
    End Select
End Sub

Public Property Get Reporter() As TReport
    If m_reporter Is Nothing Then
        Set m_reporter = New TReport
    End If
    Set Reporter = m_reporter
End Property

Public Property Let Reporter(ByVal new_reporter As TReport)
    Set Reporter = new_reporter
End Property

