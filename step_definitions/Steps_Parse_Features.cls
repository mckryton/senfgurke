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
    example_context.set_value Replace(step_expressions(1), " <br> ", vbLf), "gherkin_text"
End Sub

Public Sub Given_a_feature_F5C67A506977(example_context As TContext)
    'Given a feature
    example_context.set_value "Feature: sample feature", "gherkin_text"
End Sub

Public Sub Given_the_first_line_of_the_feature_is___wip__D41303270416(example_context As TContext)
    'Given the first line of the feature is "  @wip @important @beta"
    example_context.set_value "  @wip @important @beta" & vbLf & example_context.get_value("gherkin_text"), "gherkin_text"
End Sub

Public Sub Given_the_line_before_the_only_example_is___w_FC1ACED9ED9D(example_context As TContext)
    'Given the line before the only example is "  @wip @important @beta"
    example_context.set_value example_context.get_value("gherkin_text") & vbLf & "  @wip @important @beta" & vbLf _
                        & "Example: sample example", "gherkin_text"
End Sub

Public Sub When_the_feature_is_parsed_CEB0ED59F01F(example_context As TContext)
    'When the feature is parsed
    example_context.set_value New Collection, "parse_error"
    On Error GoTo parse_error
    example_context.set_value TFeatureParser.parse_feature(example_context.get_value("gherkin_text")), "parsed_feature"
    Exit Sub
    
parse_error:
    example_context.get_value("parse_error").Add Err.description, "err_description"
    Err.Clear
End Sub

Public Sub When_the_feature_definition_is_parsed_8C5F7D6808AE(example_context As TContext)
    'When the feature definition is parsed
    example_context.set_value New Collection, "parse_error"
    On Error GoTo parse_error
    example_context.set_value TFeatureParser.parse_feature_definition(example_context.get_value("gherkin_text")), "parsed_feature"
    Exit Sub
    
parse_error:
    example_context.get_value("parse_error").Add Err.description, "err_description"
    Err.Clear
End Sub

Public Sub Then_the_error_STR_was_raised_7C9B342540B8(example_context As TContext, step_expressions As Collection)
    'Then the error {string} was raised
    TSpec.expect(example_context.get_value("parse_error")("err_description")).to_be step_expressions(1)
End Sub

Public Sub Then_the_parsed_result_contains_a_separate_it_C01AA5CDB981(example_context As TContext)
    'Then the parsed result contains a separate item for each of the given elements
    
    Dim parsed_feature As TFeature
    Dim Rule As TRule
    Dim example As TExample
    
    Set parsed_feature = example_context.get_value("parsed_feature")
    TSpec.expect(parsed_feature.Name).to_be "sample feature"
    TSpec.expect(parsed_feature.sections.Count).to_be 1
    TSpec.expect(TypeName(parsed_feature.sections(1))).to_be "TRule"
    Set Rule = parsed_feature.sections(1)
    TSpec.expect(Rule.Name).to_be "sample rule"
    TSpec.expect(Rule.sections.Count).to_be 1
    Set example = Rule.sections(1)
    TSpec.expect(TypeName(example)).to_be "TExample"
    TSpec.expect(example.Name).to_be "sample example"
End Sub

Public Sub Then_the_parsed_result_contains_a_feature_wit_5E84028445E3(example_context As TContext, step_expressions As Collection)
    'Then the parsed result contains a feature with the name {string}
    TSpec.expect(example_context.get_value("parsed_feature").Name).to_be step_expressions(1)
End Sub

Public Sub Then_the_feature_description_is_set_to_those__3759AF1F585B(example_context As TContext)
    'Then the feature description is set to those two lines
    TSpec.expect(example_context.get_value("parsed_feature").description).to_be "this is" & vbLf & "the feature description"
End Sub

Public Sub Then_the_parsed_feature_contains_a_rule_A7C0683CBA7A(example_context As TContext)
    'Then the parsed feature contains a rule
    TSpec.expect(TypeName(example_context.get_value("parsed_feature").sections(1))).to_be "TRule"
End Sub

Public Sub Then_the_rules_description_is_set_to_STR_750959DF7853(example_context As TContext, step_expressions As Collection)
    'Then the rules description is set to {string}
     TSpec.expect(example_context.get_value("parsed_feature").sections(1).description).to_be step_expressions(1)
End Sub

Public Sub Then_the_parsed_features_contains_an_example_7B67F2AF5933(example_context As TContext)
    'Then the parsed features contains an example
    TSpec.expect(TypeName(example_context.get_value("parsed_feature").sections(1))).to_be "TExample"
    example_context.set_value example_context.get_value("parsed_feature").sections(1), "example"
End Sub

Public Sub Then_parsing_didnt_cause_any_errors_95DC5E1A0DA5(example_context As TContext)
    'Then parsing didn't cause any errors
    TSpec.expect(example_context.get_value("parse_error").Count).to_be 0
End Sub

Public Sub Then_the_parsed_result_contains_an_example_wi_5E84053FD2C3(example_context As TContext)
    'Then the parsed result contains an example with two steps
    TSpec.expect(example_context.get_value("parsed_feature").sections(1).steps.Count).to_be 2
End Sub

Public Sub Then_the_given_step_is_assigned_to_the_featur_F07FF8233D0D(example_context As TContext)
    'Then the Given step is assigned to the feature
    TSpec.expect(example_context.get_value("parsed_feature").background.steps(1).OriginalStepLine).to_be "Given one step"
End Sub

Public Sub Then_a_feature_named_STR_is_parsed_2E9486C04330(example_context As TContext, step_expressions As Collection)
    'Then a feature named {string} is parsed
    TSpec.expect(example_context.get_value("parsed_feature").Name).to_be step_expressions(1)
End Sub

Public Sub Then_no_parsed_feature_is_created_8F49EB33AFB0(example_context As TContext)
    'Then no parsed feature is created
    TSpec.expect(example_context.get_value("parsed_feature")).to_be_nothing
End Sub

Public Sub Then_a_parsed_feature_is_created_9DBF3201C55E(example_context As TContext)
    'Then a parsed feature is created
    TSpec.expect(example_context.get_value("parsed_feature")).not_to_be_nothing
End Sub

Public Sub Then_the_parsed_feature_has_INT_tags_41F92DEB0B62(example_context As TContext, step_expressions As Collection)
    'And the parsed feature has {integer} tags
    Dim parsed_feature As TFeature
    
    Set parsed_feature = example_context.get_value("parsed_feature")
    TSpec.expect(parsed_feature.tags.Count).to_be step_expressions(1)
End Sub
