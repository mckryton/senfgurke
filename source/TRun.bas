Attribute VB_Name = "TRun"
Option Explicit

Public Sub test(Optional filter_tag, Optional report_format)
    
    Dim step_definitions As Variant
    Dim features_as_text As Collection
    Dim parsed_features As Collection
    Dim parsed_feature As Variant
    Dim feature As TFeature

    If IsMissing(filter_tag) Then filter_tag = vbNullString
    If IsMissing(report_format) Then report_format = "verbose"
    Select Case report_format
        Case "verbose", "v"
            TReport.Report_Formatter = New TReportFormatterVerbose
        Case "progress", "p"
            TReport.Report_Formatter = New TReportFormatterProgress
    End Select
    Set features_as_text = TFeatureLoader.load_features
    Set parsed_features = TFeatureParser.parse_loaded_features(features_as_text)
    For Each parsed_feature In parsed_features
        Set feature = parsed_feature
        TFeatureRunner.run_feature feature, filter_tag:=filter_tag
    Next
    TReport.report_code_templates_for_missing_steps
    TConfig.StepImplementations = Nothing
End Sub

Public Sub wip()
    'wip = work in progress
    test "@wip"
End Sub
