Attribute VB_Name = "TFeatureRunner"
Option Explicit

Public Sub run_feature(feature As TFeature)
    
    Dim all_features As Variant
    Dim feature_clause As Variant
    Dim example As TExample
    Dim example_runner As TExampleRunner

    'TODO: filter tags
    Set example_runner = New TExampleRunner
    TReport.report TReport.TYPE_FEATURE_NAME, feature.Name
    TReport.report TReport.TYPE_DESC, feature.description
    For Each feature_clause In feature.Clauses
        If TypeName(feature_clause) = "TExample" Then
            Set example = feature_clause
            example_runner.run_example example
        ElseIf TypeName(feature_clause) = "TRule" Then
            TReport.report TReport.TYPE_RULE, feature_clause.WholeRule
        End If
    Next
End Sub

