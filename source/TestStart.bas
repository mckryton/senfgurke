Attribute VB_Name = "TestStart"
Option Explicit

Public Sub run_tests(Optional Tags, Optional report_format)
    
    Dim step_definitions As Variant
    Dim reporter As Variant

    If IsMissing(report_format) Then
        TReport.Report_Formatter = New TReport_Formatter_Verbose
    Else
        Debug.Print "PENDING: support for mutliple report formats"
    End If
    step_definitions = Array(New Steps_Verbose_Output, New Steps_Load_Feature_Files, New Steps_Run_Steps)
    TFeatureRunner_Old.run_features step_definitions, Tags
End Sub

Public Sub run_wip_test()
    ' wip = work in progress
    TestStart.run_tests "wip"
End Sub

