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

Public Function run_feature(feature As TFeature, Optional filter_tag, Optional silent) As Collection
    
    Dim feature_clause As Variant
    Dim example As TExample
    Dim example_runner As TExampleRunner
    Dim feature_statistics As Collection
    Dim example_statistics As Collection

    If IsMissing(filter_tag) Then filter_tag = vbNullString
    If IsMissing(silent) Then silent = False
    Set example_runner = New TExampleRunner
    example_runner.setup m_session
    Set feature_statistics = New Collection
    m_session.Reporter.report REPORT_MSG_TYPE_FEATURE_NAME, feature.Name, silent:=silent
    m_session.Reporter.report REPORT_MSG_TYPE_DESC, feature.Description, silent:=silent
    For Each feature_clause In feature.sections
        If TypeName(feature_clause) = "TExample" Then
            Set example = feature_clause
            If filter_tag = vbNullString Or ExtraVBA.collection_has_key(filter_tag, example.tags) = True Then
                If feature.Background.Steps.Count > 0 Then example.insert_background_steps feature.Background.Steps
                Set example_statistics = example_runner.run_example(example, silent:=silent)
                feature_statistics.Add example_statistics
            End If
        ElseIf TypeName(feature_clause) = "TRule" Then
            m_session.Reporter.report REPORT_MSG_TYPE_RULE, feature_clause.WholeRule, silent:=silent
        End If
    Next
    Set run_feature = feature_statistics
End Function

