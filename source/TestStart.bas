Attribute VB_Name = "TestStart"
Option Explicit

Public Sub run_tests(Optional tags, Optional report_format)
    
    Dim features As Variant
    Dim Log As Logger
    Dim Reporter As Variant

    If IsMissing(report_format) Then
        TReport.Report_Formatter = New TReport_Formatter_Verbose
    Else
        Debug.Print "PENDING: support for mutliple report formats"
    End If
    Set Log = New Logger
    features = Array(New Feature_Run_Examples, New Feature_Verbose_Output)
    TFeatureRunner.run_features features, tags
End Sub

Public Sub run_wip_test()
    ' wip = work in progress
    TestStart.run_tests "wip"
End Sub

