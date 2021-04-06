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

Public Sub Given_an_example_with_matching_step_implementation_B7CBC6F3337F(step_expressions As Collection)
    'Given an example with matching step implementations {string}
    Set m_parsed_feature = TFeatureParser.parse_feature(step_expressions(1))
End Sub

Public Sub Given_an_feature_with_one_example_with_3_steps_whe_B4401525657E()
    'Given an feature with one example with 3 steps where the 2nd step fails
    Set m_parsed_feature = TFeatureParser.parse_feature( _
                    "Feature: sample feature" & vbLf & _
                    vbLf & _
                    "  Example: sample example" & vbLf & _
                    "    Given a valid step" & vbLf & _
                    "    And an invalid step" & vbLf & _
                    "    And a valid step")
End Sub

Public Sub Given_a_valid_step_6A35DF3A18EC()
    'Given a valid step
End Sub

Public Sub Given_an_invalid_step_FF6BEF10050E()
    'And an invalid step
    TSpec.expect(1).to_be 2
End Sub

Public Sub Given_a_feature_has_a_first_example_STR_with_the_t_4BAB3B57BE7E(step_expressions As Collection)
    'Given a feature has a first example {string} with the tag {string}
    TStepVars.gherkin_text = "Feature: sample feature" & vbLf _
                        & vbLf _
                        & step_expressions(2) & vbLf _
                        & "Example: " & step_expressions(1) & vbLf _
                        & "  Given a valid step"
End Sub

Public Sub Given_the_feature_has_a_second_example_STR_without_4F58B5E86A02(step_expressions As Collection)
    'And the feature has a second example {string} without any tag
    TStepVars.gherkin_text = TStepVars.gherkin_text & vbLf _
                        & vbLf _
                        & "Example: " & step_expressions(1) & vbLf _
                        & "  Given a valid step"
End Sub

Public Sub When_the_example_is_executed_8C8F015690BC()
    'When the example is executed
    Dim feature_runner As TFeatureRunner
    
    Set feature_runner = New TFeatureRunner
    Set TStepVars.Session = New TSession
    TStepVars.Session.Reporter.enable_silent_mode
    feature_runner.setup TStepVars.Session
    Set m_feature_statistics = feature_runner.run_feature(m_parsed_feature)
End Sub

Public Sub When_the_feature_is_executed_with_tag_parameter_ST_E9DFF72092D9(step_expressions As Collection)
    'When the feature is executed with tag parameter {string}
    
    Dim parsed_feature As TFeature
    Dim feature_runner As TFeatureRunner
    
    Set feature_runner = New TFeatureRunner
    Set TStepVars.Session = New TSession
    TStepVars.Session.Reporter.enable_silent_mode
    feature_runner.setup TStepVars.Session
    Set parsed_feature = TFeatureParser.parse_feature(TStepVars.gherkin_text)
    Set m_feature_statistics = feature_runner.run_feature(parsed_feature, filter_tag:=step_expressions(1))
End Sub

Public Sub Then_step_statistics_is_STR_121A2677FC0F(step_expressions As Collection)
    'Then step statistics is {string}
    
    Dim run_statistics As Collection
    
    Set run_statistics = New Collection
    run_statistics.Add m_feature_statistics
    TSpec.expect(TStepVars.Session.get_step_statistics(run_statistics)).to_be step_expressions(1)
End Sub

Public Sub Then_only_the_example_STR_is_executed_FD3BFD3B270C(step_expressions As Collection)
    'Then only the example {string} is executed
    TSpec.expect(m_feature_statistics.Count).to_be 1
    TSpec.expect(m_feature_statistics(1)("example")).to_be "Example: " & step_expressions(1)
End Sub

