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

Public Sub Given_a_feature_STR_6A3547BB2524(example_context As TContext, step_expressions As Collection)
    'Given a feature {string}
    TStepVars.gherkin_text = Replace(step_expressions(1), " <br> ", vbLf)
End Sub

Public Sub Given_a_feature_F5C67A506977(example_context As TContext)
    'Given a feature
    TStepVars.gherkin_text = "Feature: sample feature"
End Sub

Public Sub Given_the_first_line_of_the_feature_is___wip_impor_D41303270416(example_context As TContext)
    'Given the first line of the feature is "  @wip @important @beta"
    TStepVars.gherkin_text = "  @wip @important @beta" & vbLf & TStepVars.gherkin_text
End Sub

Public Sub Given_the_line_before_the_only_example_is___wip_im_FC1ACED9ED9D(example_context As TContext)
    'Given the line before the only example is "  @wip @important @beta"
    TStepVars.gherkin_text = TStepVars.gherkin_text & vbLf & "  @wip @important @beta" & vbLf _
                        & "Example: sample example"
End Sub

Public Sub When_the_feature_is_parsed_CEB0ED59F01F(example_context As TContext)
    'When the feature is parsed
    example_context.set_value New Collection, "parse_error"
    On Error GoTo parse_error
    Set m_parsed_feature = TFeatureParser.parse_feature(TStepVars.gherkin_text)
    Exit Sub
    
parse_error:
    example_context.get_value("parse_error").Add Err.description, "err_description"
    Err.Clear
End Sub

Public Sub Then_the_error_STR_was_raised_7C9B342540B8(example_context As TContext, step_expressions As Collection)
    'Then the error {string} was raised
    TSpec.expect(example_context.get_value("parse_error")("err_description")).to_be step_expressions(1)
End Sub

Public Sub Then_the_parsed_result_contains_a_separate_item_fo_C01AA5CDB981(example_context As TContext)
    'Then the parsed result contains a separate item for each of the given elements
    
    Dim rule As TRule
    
    TSpec.expect(m_parsed_feature.name).to_be "sample feature"
    TSpec.expect(m_parsed_feature.sections.Count).to_be 1
    TSpec.expect(TypeName(m_parsed_feature.sections(1))).to_be "TRule"
    Set rule = m_parsed_feature.sections(1)
    TSpec.expect(rule.name).to_be "sample rule"
    TSpec.expect(rule.sections.Count).to_be 1
    TSpec.expect(TypeName(rule.sections(1))).to_be "TExample"
    TSpec.expect(rule.sections(1).name).to_be "sample example"
End Sub

Public Sub Then_the_parsed_result_contains_a_feature_with_the_5E84028445E3(example_context As TContext, step_expressions As Collection)
    'Then the parsed result contains a feature with the name {string}
    TSpec.expect(m_parsed_feature.name).to_be step_expressions(1)
End Sub

Public Sub Then_the_feature_description_is_set_to_those_two_l_3759AF1F585B(example_context As TContext)
    'Then the feature description is set to those two lines
    TSpec.expect(m_parsed_feature.description).to_be "this is" & vbLf & "the feature description"
End Sub

Public Sub Then_the_parsed_feature_contains_a_rule_A7C0683CBA7A(example_context As TContext)
    'Then the parsed feature contains a rule
    TSpec.expect(TypeName(m_parsed_feature.sections(1))).to_be "TRule"
End Sub

Public Sub Then_the_rules_description_is_set_to_STR_750959DF7853(example_context As TContext, step_expressions As Collection)
    'Then the rules description is set to {string}
     TSpec.expect(m_parsed_feature.sections(1).description).to_be step_expressions(1)
End Sub

Public Sub Then_the_parsed_features_contains_an_example_7B67F2AF5933(example_context As TContext)
    'Then the parsed features contains an example
    TSpec.expect(TypeName(m_parsed_feature.sections(1))).to_be "TExample"
    Set TStepVars.example = m_parsed_feature.sections(1)
End Sub

Public Sub Then_parsing_didnt_cause_any_errors_95DC5E1A0DA5(example_context As TContext)
    'Then parsing didn't cause any errors
    TSpec.expect(example_context.get_value("parse_error").Count).to_be 0
End Sub

Public Sub Then_the_parsed_result_contains_an_example_with_tw_5E84053FD2C3(example_context As TContext)
    'Then the parsed result contains an example with two steps
    TSpec.expect(m_parsed_feature.sections(1).Steps.Count).to_be 2
End Sub

Public Sub Then_the_Given_step_is_assigned_to_the_feature_9078F8233D0D(example_context As TContext)
    'Then the Given step is assigned to the feature
    TSpec.expect(m_parsed_feature.background.Steps(1).OriginalStepDefinition).to_be "Given one step"
End Sub

Public Sub Then_a_feature_named_STR_is_parsed_2E9486C04330(example_context As TContext, step_expressions As Collection)
    'Then a feature named {string} is parsed
    TSpec.expect(m_parsed_feature.name).to_be step_expressions(1)
End Sub

