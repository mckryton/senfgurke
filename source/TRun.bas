Attribute VB_Name = "TRun"
Option Explicit

Dim m_session As TSession

Public Sub test(Optional filter_tag, Optional report_format)
    
    Dim features_as_text As Collection
    Dim parsed_features As Collection
    Dim parsed_feature As Variant
    Dim feature As TFeature

    If IsMissing(filter_tag) Then filter_tag = vbNullString
    If IsMissing(report_format) Then report_format = "verbose"
    Session.set_report_format CStr(report_format)
    Set features_as_text = TFeatureLoader.load_features
    Set parsed_features = TFeatureParser.parse_loaded_features(features_as_text)
    For Each parsed_feature In parsed_features
        Set feature = parsed_feature
        TFeatureRunner.run_feature feature, filter_tag:=filter_tag
    Next
    Session.Reporter.report_code_templates_for_missing_steps
    TConfig.StepImplementations = Nothing
    Session = Nothing
End Sub

Public Sub wip()
    'wip = work in progress
    test "@wip"
End Sub

Public Property Get Session() As TSession
    If m_session Is Nothing Then
        Set m_session = New TSession
    End If
    Set Session = m_session
End Property

Public Property Let Session(ByVal new_session As TSession)
    Set m_session = new_session
End Property
