VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TFeatureRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_session As TSession
Private FeatureOrigin As String

Public Sub setup(session As TSession)
    Set m_session = session
End Sub

' run_feature runs examples inside features as well as examples inside rules
Public Sub run_feature(example_container As TExampleContainer, Optional filter_tag)
    Dim feature_section As Variant
    Dim example As TExample
    Dim step As Variant
    Dim example_runner As TExampleRunner
    Dim current_rule As TExampleContainer
    Dim current_feature As TFeature
    
    If IsMissing(filter_tag) Then filter_tag = vbNullString
    Set example_runner = New TExampleRunner
    example_runner.setup m_session, Me
    If example_container.containerType = "feature" Then
        Set current_feature = example_container
        FeatureOrigin = current_feature.FeatureFileName
        log_event EVENT_RUN_FEATURE_STARTED, current_feature.description, current_feature.name, _
                                SECTION_TYPE_FEATURE, feature_file:=current_feature.FeatureFileName
    End If
    For Each feature_section In example_container.sections
        If TypeName(feature_section) = "TExample" Then
            Set example = feature_section
            If filter_tag = vbNullString Or ExtraVBA.collection_has_key(filter_tag, example.tags) = True Then
                If Not example_container.parent Is Nothing Then
                    If example_container.parent.background.steps.Count > 0 Then
                        'if the examplecontainer is a rule then consider also the background
                        ' steps from the parent feature
                        example.insert_background_steps example_container.parent.background.steps
                    End If
                End If
                If example_container.background.steps.Count > 0 Then example.insert_background_steps example_container.background.steps
                ' if an outline table exists repeat the example for every data row in every outline table
                If example.Outlines.Count > 0 Then
                    example_runner.run_outline_example example
                Else
                    log_event EVENT_RUN_EXAMPLE_STARTED, example.description, example.OriginalHeadline, SECTION_TYPE_EXAMPLE
                    example_runner.run_example example
                End If
            End If
        ElseIf TypeName(feature_section) = "TRule" Then
            Set current_rule = feature_section
            ' a rule can contain  examples like a feature > so run it like a feature
            log_event EVENT_RUN_RULE_STARTED, feature_section.description, feature_section.wholeName, SECTION_TYPE_RULE
            run_feature current_rule, filter_tag
        End If
    Next
    If example_container.containerType = "feature" Then
        log_event EVENT_RUN_FEATURE_FINISHED, current_feature.description, current_feature.name, _
                                SECTION_TYPE_FEATURE, feature_file:=current_feature.FeatureFileName
    End If
End Sub

Public Sub log_event(event_name As String, Optional event_msg, Optional section_name, Optional section_type, _
                        Optional section_status, Optional feature_file, Optional log_time)
    If IsMissing(feature_file) Then feature_file = FeatureOrigin
    m_session.log_event event_name, event_msg, section_name, section_type, section_status, feature_file, log_time
End Sub

