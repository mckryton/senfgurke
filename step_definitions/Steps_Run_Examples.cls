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

Dim m_example As TExample
Dim m_parsed_feature As TFeature
Dim m_example_statistics As Collection

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

Public Sub When_the_example_is_executed_A6DB3176A298()
    'When the example is executed
    Dim example_runner As TExampleRunner
    
    Set example_runner = New TExampleRunner
    Set m_example_statistics = example_runner.run_example(m_parsed_feature.Clauses(1), silent:=True)
End Sub

Public Sub Then_only_two_steps_of_the_example_were_executed_73EBE317BB87()
    'Then only two steps of the example were executed

    TSpec.expect(m_example_statistics.Count).to_be 2
End Sub
