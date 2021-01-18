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

Dim m_parsed_feature As TFeature
Dim m_feature_statistics As Collection

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

Public Sub Given_a_feature_has_a_first_example_tagged_example_A6ACCC9CD3BB()
    'Given a feature has a first example "tagged example" with the tag "@sample"
    TStepVars.gherkin_text = "Feature: sample feature" & vbLf _
                        & vbLf _
                        & "@sample" & vbLf _
                        & "Example: tagged example" & vbLf _
                        & "  Given a valid step"
End Sub

Public Sub Given_the_feature_has_a_second_example_untagged_ex_EE05AE816679()
    'Given the feature has a second example "un-tagged example" without any tag
    TStepVars.gherkin_text = TStepVars.gherkin_text & vbLf _
                        & vbLf _
                        & "Example: un-tagged example" & vbLf _
                        & "  Given a valid step"
End Sub

Public Sub When_the_example_is_executed_8C8F015690BC()
    'When the example is executed
    
    Dim example_runner As TExampleRunner
    
    Set example_runner = New TExampleRunner
    Set m_feature_statistics = TFeatureRunner.run_feature(m_parsed_feature, silent:=True)
    'Set TStepVars.example_statistics = example_runner.run_example(m_parsed_feature.Clauses(1), silent:=True)
End Sub

Public Sub When_the_feature_is_executed_with_tag_parameter_sa_2662F331171B()
    'When the feature is executed with tag parameter "@sample"
    
    Dim parsed_feature As TFeature
    
    Set parsed_feature = TFeatureParser.parse_feature(TStepVars.gherkin_text)
    Set m_feature_statistics = TFeatureRunner.run_feature(parsed_feature, filter_tag:="@sample", silent:=True)
End Sub

Public Sub Then_step_statistics_is_STR_121A2677FC0F(step_expressions As Collection)
    'Then step statistics is {string}
    
    Dim run_statistics As Collection
    
    Set run_statistics = New Collection
    run_statistics.Add m_feature_statistics
    TSpec.expect(TRun.Session.get_step_statistics(run_statistics)).to_be step_expressions(1)
End Sub

Public Sub Then_only_the_example_tagged_example_is_executed_0C73E26B826F()
    'Then only the example "tagged example" is executed
    TSpec.expect(m_feature_statistics.Count).to_be 1
    TSpec.expect(m_feature_statistics(1)("example")).to_be "Example: tagged example"
End Sub
