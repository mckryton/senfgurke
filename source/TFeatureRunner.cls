VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TFeatureRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_session As TSession

Public Sub setup(Session As TSession)
    Set m_session = Session
End Sub

Public Sub run_feature(feature As TFeature, Optional filter_tag)
    
    Dim feature_clause As Variant
    Dim example As TExample
    Dim example_runner As TExampleRunner

    If IsMissing(filter_tag) Then filter_tag = vbNullString
    Set example_runner = New TExampleRunner
    example_runner.setup m_session
    m_session.Reporter.report REPORT_MSG_TYPE_FEATURE_NAME, feature.Name
    m_session.Reporter.report REPORT_MSG_TYPE_DESC, feature.Description
    For Each feature_clause In feature.sections
        If TypeName(feature_clause) = "TExample" Then
            Set example = feature_clause
            If filter_tag = vbNullString Or ExtraVBA.collection_has_key(filter_tag, example.tags) = True Then
                If feature.Background.Steps.Count > 0 Then example.insert_background_steps feature.Background.Steps
                example_runner.run_example example
            End If
        ElseIf TypeName(feature_clause) = "TRule" Then
            m_session.Reporter.report REPORT_MSG_TYPE_RULE, feature_clause.WholeRule
        End If
    Next
    m_session.statistics.log_event LOG_TYPE_FEATURE, feature.Name, ExtraVBA.get_unix_timestamp_now
End Sub

