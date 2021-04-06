VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_report_statistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_a_test_run_took_INT_ms_A7563698E569(step_expressions As Collection)
    'And a test run took {integer} ms
    TStepVars.Duration = CInt(step_expressions(1))
End Sub

Public Sub Given_one_example_in_this_run_had_INT_passed_INT_f_800581B31566(step_expressions As Collection)
    'And one example in this run had {integer} passed, {integer} failed, {integer} missing and {integer} pending steps
    Dim index As Integer
   
    Set TStepVars.RunStatistics = New Collection
    TStepVars.RunStatistics.Add New Collection          'log feature
    TStepVars.RunStatistics(1).Add New Collection       'log example
    TStepVars.RunStatistics(1)(1).Add "dummy example"   'log example name
    For index = 1 To step_expressions(1)
        TStepVars.RunStatistics(1)(1).Add Array(STATUS_OK, "dummy step name")
    Next
    For index = 1 To step_expressions(2)
        TStepVars.RunStatistics(1)(1).Add Array(STATUS_FAIL, "dummy step name")
    Next
    For index = 1 To step_expressions(3)
        TStepVars.RunStatistics(1)(1).Add Array(STATUS_MISSING, "dummy step name")
    Next
    For index = 1 To step_expressions(4)
        TStepVars.RunStatistics(1)(1).Add Array(STATUS_PENDING, "dummy step name")
    Next
End Sub

Public Sub When_the_statistics_are_reported_418C385A7E0B()
    'When the statistics are reported
    TStepVars.Session.report_statistics TStepVars.RunStatistics, TStepVars.Duration
End Sub
