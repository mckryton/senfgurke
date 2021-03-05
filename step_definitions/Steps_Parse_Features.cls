VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Parse_Features"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_a_feature_STR_6A3547BB2524(step_expressions As Collection)
    'Given a feature {string}
    TStepVars.gherkin_text = Replace(step_expressions(1), " <br> ", vbLf)
End Sub

Public Sub Given_a_feature_F5C67A506977()
    'Given a feature
    TStepVars.gherkin_text = "Feature: sample feature"
End Sub

Public Sub Given_the_first_line_of_the_feature_is___wip_impor_D41303270416()
    'Given the first line of the feature is "  @wip @important @beta"
    TStepVars.gherkin_text = "  @wip @important @beta" & vbLf & TStepVars.gherkin_text
End Sub

Public Sub Given_the_line_before_the_only_example_is___wip_im_FC1ACED9ED9D()
    'Given the line before the only example is "  @wip @important @beta"
    TStepVars.gherkin_text = TStepVars.gherkin_text & vbLf & "  @wip @important @beta" & vbLf _
                        & "Example: sample example"
End Sub

Public Sub When_the_feature_is_parsed_CEB0ED59F01F()
    'When the feature is parsed
    Set m_parsed_feature = TFeatureParser.parse_feature(TStepVars.gherkin_text)
End Sub

Public Sub Then_the_parsing_results_in_the_error_message_STR_9AC7D885B797(step_expressions As Collection)
    'Then the parsing results in the error message {string}
    TSpec.expect(m_parsed_feature.ErrorStatus).to_be step_expressions(1)
End Sub

Public Sub Then_the_parsed_result_contains_a_separate_item_fo_C01AA5CDB981()
    'Then the parsed result contains a separate item for each of the given elements
    TSpec.expect(m_parsed_feature.Name).to_be "sample feature"
    TSpec.expect(m_parsed_feature.Clauses.Count).to_be 2
    TSpec.expect(TypeName(m_parsed_feature.Clauses(1))).to_be "TRule"
    TSpec.expect(m_parsed_feature.Clauses(1).Name).to_be "sample rule"
    TSpec.expect(TypeName(m_parsed_feature.Clauses(2))).to_be "TExample"
    TSpec.expect(m_parsed_feature.Clauses(2).Name).to_be "sample example"
End Sub

Public Sub Then_the_feature_description_is_set_to_those_two_l_3759AF1F585B()
    'Then the feature description is set to those two lines
    TSpec.expect(m_parsed_feature.Description).to_be "this is" & vbLf & "the feature description"
End Sub

Public Sub Then_the_parsed_feature_contains_a_rule_A7C0683CBA7A()
    'Then the parsed feature contains a rule
    TSpec.expect(TypeName(m_parsed_feature.Clauses(1))).to_be "TRule"
End Sub

Public Sub Then_the_rules_description_is_set_to_STR_750959DF7853(step_expressions As Collection)
    'Then the rules description is set to {string}
     TSpec.expect(m_parsed_feature.Clauses(1).Description).to_be step_expressions(1)
End Sub

Public Sub Then_the_parsed_features_contains_an_example_7B67F2AF5933()
    'Then the parsed features contains an example
    TSpec.expect(TypeName(m_parsed_feature.Clauses(1))).to_be "TExample"
    Set TStepVars.example = m_parsed_feature.Clauses(1)
End Sub

Public Sub Then_the_parsed_feature_doesnt_contain_any_error_EB00DB92B133()
    'Then the parsed feature doesn't contain any error
    TSpec.expect(m_parsed_feature.ErrorStatus).to_be vbNullString
End Sub

Public Sub Then_the_parsed_result_contains_an_example_with_tw_5E84053FD2C3()
    'Then the parsed result contains an example with two steps
    TSpec.expect(m_parsed_feature.Clauses(1).Steps.Count).to_be 2
End Sub

Public Sub Then_the_Given_step_is_assigned_to_the_feature_9078F8233D0D()
    'Then the Given step is assigned to the feature
    TSpec.expect(m_parsed_feature.Background.Steps(1).OriginalStepDefinition).to_be "Given one step"
End Sub

