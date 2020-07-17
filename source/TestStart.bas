Attribute VB_Name = "TestStart"
Option Explicit

Public Sub run_tests(Optional tags, Optional report_format)
    
    Dim feature_runner As TFeatureRunner
    Dim features As Variant
    Dim log As Logger
    Dim Reporter As Variant

    If IsMissing(report_format) Then
        TReport.Report_Formatter = New TReport_Formatter_Verbose
    Else
        Debug.Print "PENDING: support for mutliple report formats"
    End If
    Set log = New Logger
    features = Array(New Feature_Execute_Examples, New Feature_Verbose_Output)
    Set feature_runner = New TFeatureRunner
    feature_runner.run_features features, tags
End Sub

Public Sub run_wip_test()
    ' wip = work in progress
    TestStart.run_tests "wip"
End Sub

