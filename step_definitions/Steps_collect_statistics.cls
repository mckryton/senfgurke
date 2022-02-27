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

Public Sub Given_a_new_test_run_started_collecting_statistics_152BAB44CFC9(example_context As TContext)
    'Given a new test run started collecting statistics
    Dim session As TSession
    Dim event_stream_listener As EventListener
    Dim step_implementations As Collection
    
    Set step_implementations = New Collection
    step_implementations.Add New Steps_predefined_steps
    Set session = THelper.new_TSession
    example_context.set_value session, "session"
    session.reporter.enable_silent_mode
    session.StepImplementations = step_implementations
End Sub

Public Sub Given_duration_of_a_test_run_is_INT_ms_CB357F7BA1F9(example_context As TContext, step_expressions As Collection)
    'Given duration of a test run is {integer} ms
    example_context.set_value step_expressions(1), "duration"
End Sub

Public Sub When_the_feature_is_executed_and_statistics_are_co_0BF6FCF4961B(example_context As TContext)
    'When the feature is executed and statistics are collected
    Dim Feature As TFeature
    Dim parsed_features As Collection
    Dim session As TSession
    
    Set session = example_context.get_value("session")
    Set Feature = TFeatureParser.parse_feature(example_context.get_value("gherkin_text"))
    Set parsed_features = New Collection
    parsed_features.Add Feature
    session.run_parsed_features parsed_features
End Sub

Public Sub When_the_duration_is_calculated_4FCC027398DA(example_context As TContext)
    'When the duration is calculated
    example_context.set_value example_context.get_value("session").get_duration_as_text(example_context.get_value("duration")), "report_output"
End Sub

Public Sub When_the_results_for_the_steps_are_summed_up_EE8E882F4756(example_context As TContext)
    'When the results for the steps are summed up
    example_context.set_value example_context.get_value("session").statistics.get_statistics_summary, "statistic_results"
End Sub

Public Sub Then_start_and_end_time_for_the_test_run_are_set_262A8F19D7C7(example_context As TContext)
    'Then start and end time for the test run are set
    TSpec.expect(CLng(example_context.get_value("session").StartTime)).not_to_be 0
    TSpec.expect(CLng(example_context.get_value("session").EndTime)).not_to_be 0
End Sub

Public Sub Then_the_resulting_output_is_STR_8EBDD0FB2136(example_context As TContext, step_expressions As Collection)
    'Then the resulting output is {string}
    TSpec.expect(example_context.get_value("report_output")).to_be step_expressions(1)
End Sub

Public Sub Then_statistics_results_for_steps_is_STR_305CDE5C3438(example_context As TContext, step_expressions As Collection)
    'Then statistics results for steps is {string}
    TSpec.expect(example_context.get_value("statistic_results")).to_be step_expressions(1)
End Sub

