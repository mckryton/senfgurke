Attribute VB_Name = "TestStart"
Option Explicit

Public Sub run_acceptance_tests(Optional pTags)
    
    Dim case_runner As TCaseRunner
    Dim acceptance_testcases As Variant
    Dim Log As logger

    On Error GoTo error_handler
    Set Log = New logger
    acceptance_testcases = Array(New Feature_ApplyRules, New Feature_Rule_Permitted_Fonts)
    Set case_runner = New TCaseRunner
    case_runner.run_testcases acceptance_testcases, pTags
    Exit Sub

error_handler:
    Log.log_function_error "TestStart.run_acceptance_tests"
End Sub

Public Sub run_acceptance_wip_tests()
    'wip = work in progress
    TestStart.run_acceptance_tests "wip"
End Sub


Public Sub run_unit_tests(Optional pTags)
    
    Dim case_runner As TCaseRunner
    Dim unit_testcases As Variant
    Dim Log As logger

    On Error GoTo error_handler
    Set Log = New logger
    unit_testcases = Array(New Unit_ReadConfig, New Unit_ChooseTargetPresentation, New Unit_SetupRules)
    Set case_runner = New TCaseRunner
    case_runner.run_testcases unit_testcases, pTags
    Exit Sub

error_handler:
    Log.log_function_error "TestStart.run_unit_tests"
End Sub


Public Sub run_unit_wip_tests()
    'wip = work in progress
    TestStart.run_unit_tests "wip"
End Sub

