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

Public Sub run_feature(example_container As TExampleContainer, Optional filter_tag)
    
    Dim feature_section As Variant
    Dim example As TExample
    Dim example_runner As TExampleRunner
    Dim current_rule As TExampleContainer

    If IsMissing(filter_tag) Then filter_tag = vbNullString
    Set example_runner = New TExampleRunner
    example_runner.setup m_session
    If TypeName(feature_section) = "TFeature" Then
        m_session.Reporter.report REPORT_MSG_TYPE_FEATURE_NAME, example_container.name
        m_session.Reporter.report REPORT_MSG_TYPE_DESC, example_container.description
    Else
        m_session.Reporter.report REPORT_MSG_TYPE_RULE, example_container.wholeName
    End If
    For Each feature_section In example_container.sections
        If TypeName(feature_section) = "TExample" Then
            Set example = feature_section
            If filter_tag = vbNullString Or ExtraVBA.collection_has_key(filter_tag, example.tags) = True Then
                If Not example_container.parent Is Nothing Then
                    If example_container.parent.background.Steps.Count > 0 Then
                        'if the examplecontainer is a rule then consider also the background
                        ' steps from the parent feature
                        example.insert_background_steps example_container.parent.background.Steps
                    End If
                End If
                If example_container.background.Steps.Count > 0 Then example.insert_background_steps example_container.background.Steps
                example_runner.run_example example
            End If
        ElseIf TypeName(feature_section) = "TRule" Then
            Set current_rule = feature_section
            ' a rule can contain  examples like a feature > so run it like a feature
            run_feature current_rule, filter_tag
        End If
    Next
    If TypeName(feature_section) = "TFeature" Then
        m_session.statistics.log_event LOG_TYPE_FEATURE, example_container.wholeName, ExtraVBA.get_unix_timestamp_now
    Else
        m_session.statistics.log_event LOG_TYPE_RULE, example_container.wholeName, ExtraVBA.get_unix_timestamp_now
    End If
End Sub

