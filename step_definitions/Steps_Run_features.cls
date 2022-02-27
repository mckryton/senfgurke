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
    Dim event_listener As EventListener
    Dim parsed_feature As TFeature
    Dim feature_runner As TFeatureRunner
    Dim session As TSession
    
    Set feature_runner = THelper.new_TFeatureRunner
    Set event_listener = New EventListener
    Set session = THelper.new_TSession
    Set event_listener.SenfgurkeEventLogger = session.EventLogger
    session.reporter.enable_silent_mode
    feature_runner.setup session
    Set parsed_feature = TFeatureParser.parse_feature(example_context.get_value("gherkin_text"))
    feature_runner.run_feature parsed_feature
    example_context.set_value event_listener, "event_listener"
End Sub

Public Sub When_the_feature_is_executed_with_tag_parameter_ST_E9DFF72092D9(example_context As TContext, step_expressions As Collection)
    'When the feature is executed with tag parameter {string}
    Dim parsed_feature As TFeature
    Dim feature_runner As TFeatureRunner
    Dim session As TSession
    Dim event_listener As EventListener
    
    Set feature_runner = THelper.new_TFeatureRunner
    Set event_listener = New EventListener
    Set session = THelper.new_TSession
    Set event_listener.SenfgurkeEventLogger = session.EventLogger
    session.reporter.enable_silent_mode
    feature_runner.setup session
    Set parsed_feature = TFeatureParser.parse_feature(example_context.get_value("gherkin_text"))
    feature_runner.run_feature parsed_feature, filter_tag:=step_expressions(1)
    example_context.set_value event_listener, "event_listener"
End Sub

Public Sub Then_the_background_step_is_executed_once_before_e_AD9FFABBC1CE(example_context As TContext)
    'Then the background step is executed once before each example
    Dim event_listener As EventListener
    Dim step_events As Collection
    Dim log_event As Variant
    
    Set event_listener = example_context.get_value("event_listener")
    Set step_events = event_listener.filter_events(EVENT_RUN_STEP_FINISHED)
    TSpec.expect(step_events.Count).to_be 4
    TSpec.expect(step_events(1).SectionName).includes_text "background step"
    TSpec.expect(step_events(3).SectionName).includes_text "background step"
End Sub

Public Sub Then_only_the_example_STR_is_executed_FD3BFD3B270C(example_context As TContext, step_expressions As Collection)
    'Then only the example {string} is executed
    Dim event_listener As EventListener
    Dim example_events As Collection
    Dim log_event As Variant
    
    Set event_listener = example_context.get_value("event_listener")
    Set example_events = event_listener.filter_events(EVENT_RUN_EXAMPLE_FINISHED)
    TSpec.expect(example_events.Count).to_be 1
    TSpec.expect(example_events(1).SectionName).includes_text "tagged example"
End Sub
