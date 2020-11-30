Attribute VB_Name = "TFeatureRunner"
Option Explicit

Public Function run_feature(feature As TFeature, Optional filter_tag, Optional silent) As Collection
    
    Dim all_features As Variant
    Dim feature_clause As Variant
    Dim example As TExample
    Dim example_runner As TExampleRunner
    Dim feature_statistics As Collection
    Dim example_statistics As Collection

    If IsMissing(filter_tag) Then filter_tag = vbNullString
    If IsMissing(silent) Then silent = False
    Set example_runner = New TExampleRunner
    Set feature_statistics = New Collection
    If Not silent Then
        TReport.report TReport.TYPE_FEATURE_NAME, feature.Name
        TReport.report TReport.TYPE_DESC, feature.Description
    End If
    For Each feature_clause In feature.Clauses
        If TypeName(feature_clause) = "TExample" Then
            Set example = feature_clause
            If filter_tag = vbNullString Or ExtraVBA.collection_has_key(filter_tag, example.Tags) = True Then
                If feature.Background.Steps.Count > 0 Then example.insert_background_steps feature.Background.Steps
                Set example_statistics = example_runner.run_example(example, silent:=silent)
                feature_statistics.Add example_statistics
            End If
        ElseIf TypeName(feature_clause) = "TRule" Then
            TReport.report TReport.TYPE_RULE, feature_clause.WholeRule
        End If
    Next
    Set run_feature = feature_statistics
End Function

