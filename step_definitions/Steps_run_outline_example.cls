VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_run_outline_example"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_an_example_for_an_outline_STR_9C2B5479F82D(example_context As TContext, step_expressions As Collection)
    'Given an example for an outline
    Dim parsed_feature As TFeature
    Dim gherkin As String
    
    'to be executed examples must be part of a feature
    gherkin = "Feature: sample" & vbLf & step_expressions(1)
    Set parsed_feature = TFeatureParser.parse_feature(gherkin)
    example_context.set_value parsed_feature, "parsed_feature"
End Sub

Public Sub Then_in_total_INT_steps_where_executed_5C9B02007642(example_context As TContext, step_expressions As Collection)
    'Then in total {integer} steps where executed
    Dim event_listener As EventListener
    Dim event_list As Collection
    
    Set event_listener = example_context.get_value("event_listener")
    Set event_list = event_listener.filter_events(EVENT_RUN_STEP_FINISHED)
    tspec.expect(event_list.Count).to_be step_expressions(1)
End Sub

Public Sub Then_the_step_definition_has_changed_to_STR_9017F65FCB1B(example_context As TContext, step_expressions As Collection)
    'Then the step definition has changed to {string}
    Dim event_listener As EventListener
    Dim event_list As Collection
    
    Set event_listener = example_context.get_value("event_listener")
    Set event_list = event_listener.filter_events(EVENT_RUN_STEP_FINISHED)
    tspec.expect(event_list(1).SectionName).to_be step_expressions(1)
End Sub

Public Sub Then_step_definitions_should_have_changed_to__407185A0FF0E(example_context As TContext, step_expressions As Collection)
    'Then step definitions should have changed to {string}
    Dim event_listener As EventListener
    Dim event_list As Collection
    Dim run_step_event As TEvent
    Dim step_definitions As String
    
    step_definitions = vbNullString
    Set event_listener = example_context.get_value("event_listener")
    Set event_list = event_listener.filter_events(EVENT_RUN_STEP_FINISHED)
    For Each run_step_event In event_list
        If step_definitions <> vbNullString Then step_definitions = step_definitions & vbLf
        step_definitions = step_definitions & run_step_event.SectionName
    Next
    tspec.expect(step_definitions).to_be step_expressions(1)
End Sub
