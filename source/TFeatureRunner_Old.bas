Attribute VB_Name = "TFeatureRunner_Old"
Option Explicit

Public Sub run_features(features As Variant, Optional pTags)
    
    Dim all_features As Variant
    Dim feature As Variant
    Dim should_test_run As Boolean
    Dim examples_list As Collection
    Dim example As Variant

    For Each feature In features
        should_test_run = True
        If Not IsMissing(pTags) Then
            should_test_run = test_has_tag(pTags, feature)
        End If
        If should_test_run Then
            TReport.report TReport.TYPE_FEATURE_NAME, TypeName(feature)
            TReport.report TReport.TYPE_DESC, feature.Description
            Set examples_list = feature.Examples
            For Each example In examples_list
                TExampleRunner_Old.run_example example, feature
            Next
            Set feature = Nothing
        End If
    Next
End Sub

Private Function test_has_tag(pTags As Variant, pTesTFeature As Variant)

    Dim input_tags As Variant
    Dim feature_tags As Variant
    Dim match As Variant
    Dim tag As Variant
    
    test_has_tag = False
    input_tags = Split(Replace(pTags, " ", ""), ",")
    feature_tags = Split(Replace(pTesTFeature.tags, " ", ""), ",")
    For Each tag In input_tags
        match = Filter(feature_tags, tag, True, vbBinaryCompare)
        If UBound(match) > -1 Then
            test_has_tag = True
        End If
    Next
End Function
