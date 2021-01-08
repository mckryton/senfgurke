VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_collect_statistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_a_feature_description_with_a_step_9DB173A5EA30()
    'Given a feature description with a step
    Set TStepVars.ParsedFeatures = New Collection
    TStepVars.ParsedFeatures.Add TFeatureParser.parse_feature("Feature: sample" & vbLf _
                                                                & "Example: sample" & vbLf _
                                                                & "Given a step")
End Sub

Public Sub Given_duration_of_a_test_run_is_INT_ms_CB357F7BA1F9(step_expressions As Collection)
    'Given duration of a test run is {integer} ms
    TStepVars.Duration = step_expressions(1)
End Sub

Public Sub When_the_feature_is_executed_and_statistics_are_co_0BF6FCF4961B()
    'When the feature is executed and statistics are collected
    Set TStepVars.Session = New TSession
    TStepVars.Session.run_features TStepVars.ParsedFeatures, silent:=True
End Sub

Public Sub When_the_duration_is_calculated_4FCC027398DA()
    'When the duration is calculated
    Set TStepVars.Session = New TSession
    TStepVars.ReportOutput = TStepVars.Session.get_duration_in_ms(0, TStepVars.Duration)
End Sub

Public Sub Then_start_and_end_time_for_the_test_run_are_set_262A8F19D7C7()
    'Then start and end time for the test run are set
    TSpec.expect(CLng(TStepVars.Session.StartTime)).not_to_be 0
    TSpec.expect(CLng(TStepVars.Session.EndTime)).not_to_be 0
End Sub




