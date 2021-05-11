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

Private m_parsed_feature As TFeature
Private m_feature_statistics As Collection

Public Sub Given_an_example_with_matching_step_implementation_B7CBC6F3337F(context As TContext, step_expressions As Collection)
    'Given an example with matching step implementations {string}
    Set m_parsed_feature = TFeatureParser.parse_feature(step_expressions(1))
End Sub

Public Sub Given_an_feature_with_one_example_with_3_steps_whe_B4401525657E(context As TContext)
    'Given an feature with one example with 3 steps where the 2nd step fails
    Set m_parsed_feature = TFeatureParser.parse_feature( _
                    "Feature: sample feature" & vbLf & _
                    vbLf & _
                    "  Example: sample example" & vbLf & _
                    "    Given a valid step" & vbLf & _
                    "    And an invalid step" & vbLf & _
                    "    And a valid step")
End Sub

Public Sub Given_a_valid_step_6A35DF3A18EC(context As TContext)
    'Given a valid step
End Sub

Public Sub Given_an_invalid_step_FF6BEF10050E(context As TContext)
    'And an invalid step
    TSpec.expect(1).to_be 2
End Sub

Public Sub When_the_example_is_executed_8C8F015690BC(context As TContext)
    'When the example is executed
    Dim feature_runner As TFeatureRunner
    
    Set feature_runner = New TFeatureRunner
    Set TStepVars.session = New TSession
    TStepVars.session.Reporter.enable_silent_mode
    feature_runner.setup TStepVars.session
    feature_runner.run_feature m_parsed_feature
End Sub

Public Sub Then_step_statistics_is_STR_121A2677FC0F(context As TContext, step_expressions As Collection)
    'Then step statistics is {string}
    TSpec.expect(TStepVars.session.statistics.get_statistics_summary).to_be step_expressions(1)
End Sub


