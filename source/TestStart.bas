Attribute VB_Name = "TestStart"
Option Explicit

Public Sub run_tests(Optional pTags)
    
    Dim feature_runner As TFeatureRunner
    Dim features As Variant
    Dim Log As Logger

    Set Log = New Logger
    features = Array(New Feature_Execute_Examples, New Feature_Verbose_Output)
    Set feature_runner = New TFeatureRunner
    feature_runner.run_features features, pTags
End Sub

Public Sub run_wip_test()
    ' wip = work in progress
    TestStart.run_tests "wip"
End Sub

