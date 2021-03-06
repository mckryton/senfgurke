VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Run_features"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub When_the_feature_is_executed_8C8F454A90BC(example_context As TContext)
    'When the feature is executed
    
    Dim parsed_feature As TFeature
    Dim feature_runner As TFeatureRunner
    Dim session As TSession
    
    Set feature_runner = New TFeatureRunner
    Set session = New TSession
    example_context.set_value session, "session"
    session.Reporter.enable_silent_mode
    session.statistics.enable_event_log
    feature_runner.setup session
    Set parsed_feature = TFeatureParser.parse_feature(example_context.get_value("gherkin_text"))
    feature_runner.run_feature parsed_feature
End Sub

Public Sub When_the_feature_is_executed_with_tag_parameter_ST_E9DFF72092D9(example_context As TContext, step_expressions As Collection)
    'When the feature is executed with tag parameter {string}
    
    Dim parsed_feature As TFeature
    Dim feature_runner As TFeatureRunner
    Dim session As TSession
    
    Set feature_runner = New TFeatureRunner
    Set session = New TSession
    example_context.set_value session, "session"
    session.Reporter.enable_silent_mode
    session.statistics.enable_event_log
    feature_runner.setup session
    Set parsed_feature = TFeatureParser.parse_feature(example_context.get_value("gherkin_text"))
    feature_runner.run_feature parsed_feature, filter_tag:=step_expressions(1)
End Sub

Public Sub Then_the_background_step_is_executed_once_before_e_AD9FFABBC1CE(example_context As TContext)
    'Then the background step is executed once before each example
    Dim step_events As Collection
    Dim log_event As Variant
    
    Set step_events = New Collection
    For Each log_event In example_context.get_value("session").statistics.event_log
        If log_event("type") = LOG_EVENT_RUN_STEP Then step_events.Add log_event
    Next
    TSpec.expect(step_events.Count).to_be 4
    TSpec.expect(step_events(1)("msg")).includes_text "background step"
    TSpec.expect(step_events(3)("msg")).includes_text "background step"
End Sub

Public Sub Then_only_the_example_STR_is_executed_FD3BFD3B270C(example_context As TContext, step_expressions As Collection)
    'Then only the example {string} is executed
    Dim example_events As Collection
    Dim log_event As Variant
    
    Set example_events = New Collection
    For Each log_event In example_context.get_value("session").statistics.event_log
        If log_event("type") = LOG_EVENT_RUN_EXAMPLE Then example_events.Add log_event
    Next
    TSpec.expect(example_events.Count).to_be 1
    TSpec.expect(example_events(1)("msg")).includes_text "tagged example"
End Sub
