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

Public Sub Given_a_feature_has_a_first_example_tagged_example_A6ACCC9CD3BB()
    'Given a feature has a first example "tagged example" with the tag "@sample"
    TStepVars.feature_text = "Feature: sample feature" & vbLf _
                        & vbLf _
                        & "@sample" & vbLf _
                        & "Example: tagged example" & vbLf _
                        & "  Given a valid step"
End Sub

Public Sub Given_the_feature_has_a_second_example_untagged_ex_EE05AE816679()
    'Given the feature has a second example "un-tagged example" without any tag
    TStepVars.feature_text = TStepVars.feature_text & vbLf _
                        & vbLf _
                        & "Example: un-tagged example" & vbLf _
                        & "  Given a valid step"
End Sub

Public Sub When_the_example_is_executed_A6DB3176A298()
    'When the example is executed
    
    Dim example_runner As TExampleRunner
    
    Set example_runner = New TExampleRunner
    Set TStepVars.example_statistics = example_runner.run_example(m_parsed_feature.Clauses(1), silent:=True)
End Sub

Public Sub When_the_feature_is_executed_with_tag_parameter_sa_2662F331171B()
    'When the feature is executed with tag parameter "@sample"
    
    Dim parsed_feature As TFeature
    
    Set parsed_feature = TFeatureParser.parse_feature(TStepVars.feature_text)
    Set m_feature_statistics = TFeatureRunner.run_feature(parsed_feature, filter_tag:="@sample", silent:=True)
End Sub

Public Sub Then_only_two_steps_of_the_example_were_executed_73EBE317BB87()
    'Then only two steps of the example were executed
    TSpec.expect(TStepVars.example_statistics.Count).to_be 3 'stats count includes example name
End Sub

Public Sub Then_only_the_example_tagged_example_is_executed_0C73E26B826F()
    'Then only the example "tagged example" is executed
    TSpec.expect(m_feature_statistics.Count).to_be 1
    TSpec.expect(m_feature_statistics(1)("example")).to_be "Example: tagged example"
End Sub
