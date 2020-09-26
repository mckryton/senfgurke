Attribute VB_Name = "TFeatureRunner"
Option Explicit

Public Sub run_feature(feature As TFeature)
    
    Dim all_features As Variant
    Dim feature_clause As Variant
    Dim example As TExample

    'TODO: filter tags
    TReport.report TReport.TYPE_FEATURE_NAME, feature.Name
    TReport.report TReport.TYPE_DESC, feature.Description
    For Each feature_clause In feature.Clauses
        If TypeName(feature_clause) = "TExample" Then
            Set example = feature_clause
            TExampleRunner.run_example example
        End If
    Next
End Sub

