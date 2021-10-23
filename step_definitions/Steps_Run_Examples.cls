VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Run_Examples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_an_example_with_matching_step_implementation_B7CBC6F3337F(example_context As TContext, step_expressions As Collection)
    'Given an example with matching step implementations {string}
    example_context.set_value TFeatureParser.parse_feature(step_expressions(1)), "parsed_feature"
End Sub

Public Sub Given_an_feature_with_one_example_with_3_steps_whe_B4401525657E(example_context As TContext)
    'Given an feature with one example with 3 steps where the 2nd step fails
    example_context.set_value TFeatureParser.parse_feature( _
                    "Feature: sample feature" & vbLf & _
                    vbLf & _
                    "  Example: sample example" & vbLf & _
                    "    Given a valid step" & vbLf & _
                    "    And an invalid step" & vbLf & _
                    "    And a valid step"), _
                    "parsed_feature"
End Sub

Public Sub When_the_example_is_executed_8C8F015690BC(example_context As TContext)
    'When the example is executed
    Dim feature_runner As TFeatureRunner
    Dim parsed_feature As TFeature
    Dim session As TSession
    Dim step_classes As Collection
    
    Set step_classes = New Collection
    step_classes.Add New Steps_predefined_steps
    Set feature_runner = THelper.new_TFeatureRunner
    If Not example_context.value_exists("parsed_feature") Then
        Set parsed_feature = TFeatureParser.parse_feature(example_context.get_value("gherkin_text"))
    Else
        Set parsed_feature = example_context.get_value("parsed_feature")
    End If
    Set session = THelper.new_TSession
    session.StepImplementations = step_classes
    session.reporter.enable_silent_mode
    example_context.set_value session, "session"
    feature_runner.setup session
    feature_runner.run_feature parsed_feature
End Sub

Public Sub Then_step_statistics_is_STR_121A2677FC0F(example_context As TContext, step_expressions As Collection)
    'Then step statistics is {string}
    TSpec.expect(example_context.get_value("session").statistics.get_statistics_summary).to_be step_expressions(1)
End Sub


