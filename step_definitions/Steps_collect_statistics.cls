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

Public Sub Given_a_new_test_run_started_collecting_statistics_152BAB44CFC9()
    'Given a new test run started collecting statistics
    Set TStepVars.Session = New TSession
    TStepVars.Session.Reporter.enable_silent_mode
End Sub

Public Sub Given_a_feature_description_with_a_step_9DB173A5EA30()
    'Given a feature description with a step
    Set TStepVars.ParsedFeatures = New Collection
    TStepVars.ParsedFeatures.Add TFeatureParser.parse_feature("Feature: sample" & vbLf _
                                                                & "Example: sample" & vbLf _
                                                                & "Given a missing step")
End Sub

Public Sub Given_duration_of_a_test_run_is_INT_ms_CB357F7BA1F9(step_expressions As Collection)
    'Given duration of a test run is {integer} ms
    TStepVars.Duration = step_expressions(1)
End Sub

Public Sub Given_a_feature_with_an_example_with_a_missing_ste_63C4BE7D2E08()
    'Given a feature with an example with a missing step
    Set TStepVars.ParsedFeatures = New Collection
    TStepVars.ParsedFeatures.Add TFeatureParser.parse_feature("Feature: sample" & vbLf _
                                                                & "Example: sample" & vbLf _
                                                                & "Given a missing step")
End Sub

Public Sub Given_an_example_with_INT_steps_was_executed_succe_B7CBEFE9F6DC(step_expressions As Collection)
    'Given an example with {integer} steps was executed successful
   
    Dim index As Integer
   
    For index = 1 To step_expressions(1)
        TStepVars.Session.statistics.count_event LOG_TYPE_STEP, "OK"
    Next
End Sub

Public Sub When_the_feature_is_executed_and_statistics_are_co_0BF6FCF4961B()
    'When the feature is executed and statistics are collected
    TStepVars.Session.run_features TStepVars.ParsedFeatures
End Sub

Public Sub When_the_duration_is_calculated_4FCC027398DA()
    'When the duration is calculated
    TStepVars.ReportOutput = TStepVars.Session.get_duration_as_text(TStepVars.Duration)
End Sub

Public Sub When_the_results_for_the_steps_are_summed_up_EE8E882F4756()
    'When the results for the steps are summed up
    TStepVars.StatisticResults = TStepVars.Session.statistics.get_statistics_summary
End Sub

Public Sub Then_start_and_end_time_for_the_test_run_are_set_262A8F19D7C7()
    'Then start and end time for the test run are set
    TSpec.expect(CLng(TStepVars.Session.StartTime)).not_to_be 0
    TSpec.expect(CLng(TStepVars.Session.EndTime)).not_to_be 0
End Sub

Public Sub Then_the_resulting_output_is_STR_8EBDD0FB2136(step_expressions As Collection)
    'Then the resulting output is {string}
    TSpec.expect(TStepVars.ReportOutput).to_be step_expressions(1)
End Sub

Public Sub Then_one_step_with_its_result_was_counted_69B81E54CF8B()
    'Then one step with it's result was counted
    TSpec.expect(TStepVars.Session.statistics.step_statistics("MISSING")).to_be 1
End Sub

Public Sub Then_statistics_results_for_steps_is_STR_305CDE5C3438(step_expressions As Collection)
    'Then statistics results for steps is {string}
    TSpec.expect(TStepVars.StatisticResults).to_be step_expressions(1)
End Sub
