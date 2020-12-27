Attribute VB_Name = "TRun"
Option Explicit

Dim m_logger As Logger
Dim m_reporter As TReport

Public Sub test(Optional filter_tag, Optional report_format)
    
    Dim step_definitions As Variant
    Dim features_as_text As Collection
    Dim parsed_features As Collection
    Dim parsed_feature As Variant
    Dim feature As TFeature

    Set Reporter = New TReport
    If IsMissing(filter_tag) Then filter_tag = vbNullString
    If IsMissing(report_format) Then report_format = "verbose"
    Select Case report_format
        Case "verbose", "v"
            Reporter.Report_Formatter = New TReportFormatterVerbose
        Case "progress", "p"
            Reporter.Report_Formatter = New TReportFormatterProgress
    End Select
    Set features_as_text = TFeatureLoader.load_features
    Set parsed_features = TFeatureParser.parse_loaded_features(features_as_text)
    For Each parsed_feature In parsed_features
        Set feature = parsed_feature
        TFeatureRunner.run_feature feature, filter_tag:=filter_tag
    Next
    Reporter.report_code_templates_for_missing_steps
    TConfig.StepImplementations = Nothing
    Reporter = Nothing
End Sub

Public Sub wip()
    'wip = work in progress
    test "@wip"
End Sub

Public Property Get Log() As Logger

    If m_logger Is Nothing Then
        Set m_logger = New Logger
    End If
    Set Log = m_logger
End Property

Public Property Get Reporter() As TReport
    Set Reporter = m_reporter
End Property

Public Property Let Reporter(ByVal new_reporter As TReport)
    Set m_reporter = new_reporter
End Property
