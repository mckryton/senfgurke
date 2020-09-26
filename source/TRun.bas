Attribute VB_Name = "TRun"
Option Explicit

Public Sub test(Optional tags, Optional report_format)
    
    Dim step_definitions As Variant
    Dim features_as_text As Collection
    Dim parsed_features As Collection
    Dim parsed_feature As Variant
    Dim feature As TFeature

    If IsMissing(report_format) Then
        TReport.Report_Formatter = New TReport_Formatter_Verbose
    Else
        Debug.Print "PENDING: support for mutliple report formats"
    End If
    Set features_as_text = TFeatureLoader.load_features
    Set parsed_features = TFeatureParser.parse_feature_list(features_as_text)
    For Each parsed_feature In parsed_features
        Set feature = parsed_feature
        TFeatureRunner.run_feature feature
    Next
End Sub
