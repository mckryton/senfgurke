Attribute VB_Name = "TestStart"
Option Explicit

Public Sub run_tests(Optional pTags)
    
    Dim feature_runner As TFeatureRunner
    Dim acceptance_testcases As Variant
    Dim Log As Logger

    On Error GoTo error_handler
    Set Log = New Logger
    acceptance_testcases = Array(New Feature_ApplyRules)
    Set feature_runner = New TFeatureRunner
    feature_runner.run_testcases acceptance_testcases, pTags
End Sub

Public Sub run_acceptance_tests()
    TestStart.run_tests "feature"
End Sub

Public Sub run_acceptance_wip_tests()
    'wip = work in progress
    TestStart.run_tests "feature,wip"
End Sub

Public Sub run_unit_tests()
    TestStart.run_tests "unit"
End Sub

Public Sub run_unit_wip_tests()
    'wip = work in progress
    TestStart.run_tests "unit,wip"
End Sub

