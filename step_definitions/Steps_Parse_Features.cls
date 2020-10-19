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

Public Sub Given_a_feature_this_is_just_some_text_br_without__E5818F3BFB1B()
    'Given a feature "this is just some text <br> without a feature keyword"
    TStepVars.feature_text = Replace("this is just some text <br> without a feature keyword", " <br> ", vbLf)
End Sub

Public Sub Given_a_feature_named_sample_feature_FD3D0FF62C49()
    'Given a feature named "sample feature"
    TStepVars.feature_text = "Feature: sample feature"
End Sub

Public Sub Given_the_feature_includes_a_rule_this_is_a_sample_1BF5A1CEA9A0()
    'Given the feature includes a rule "this is a sample rule"
    TStepVars.feature_text = TStepVars.feature_text & vbLf & "Rule: this is a sample rule"
End Sub

Public Sub Given_the_feature_includes_an_example_this_is_an_e_1BF5BD6536EF()
    'Given the feature includes an example "this is an example"
    TStepVars.feature_text = TStepVars.feature_text & vbLf & "Example: this is an example"
End Sub

Public Sub Given_a_feature_starting_with___Feature_sample_fea_C087A1013C38()
    'Given a feature starting with "  Feature: sample feature "
    TStepVars.feature_text = "  Feature: sample feature "
End Sub

Public Sub Given_the_feature_continues_with_the_line_____Rule_40058483AAFA()
    'Given the feature continues with the line "    Rule: this is a rule"
    TStepVars.feature_text = TStepVars.feature_text & vbLf & "    Rule: this is a rule"
End Sub

Public Sub Given_the_line_with_the_feature_keyword_is_followe_EB6ACC822715()
    'Given the line with the feature keyword is followed by two lines "  this is<br>  the description"
    TStepVars.feature_text = TStepVars.feature_text & vbLf & "  this is" & vbLf & "  the description"
End Sub

Public Sub Given_the_rule_is_followed_by_a_line___this_is_a_d_CCBFA097ACF8()
    'Given the rule is followed by a line "  this is a description"
    TStepVars.feature_text = TStepVars.feature_text & vbLf & "  this is a description"
End Sub

Public Sub Given_a_feature_F5C67A506977()
    'Given a feature
    TStepVars.feature_text = "Feature: sample feature"
End Sub

Public Sub Given_a_feature_tag1_tag2brFeature_sample_feature_015C96EDA874()
    'Given a feature "@tag1 @tag2<br>Feature: sample feature"
    TStepVars.feature_text = "@tag1 @tag2" & vbLf _
                        & "Feature: sample feature"
End Sub

Public Sub Given_the_first_line_of_the_feature_is___wip_impor_D41303270416()
    'Given the first line of the feature is "  @wip @important @beta"
    TStepVars.feature_text = "  @wip @important @beta" & vbLf & TStepVars.feature_text
End Sub

Public Sub Given_the_line_before_the_only_example_is___wip_im_FC1ACED9ED9D()
    'Given the line before the only example is "  @wip @important @beta"
    TStepVars.feature_text = TStepVars.feature_text & vbLf & "  @wip @important @beta" & vbLf _
                        & "Example: sample example"
End Sub

Public Sub Given_the_first_line_of_the_feature_is__this_is_a__D211BCD22A79()
    'Given the first line of the feature is "# this is a comment"
    TStepVars.feature_text = "# this is a comment" & vbLf & TStepVars.feature_text
End Sub

Public Sub Given_the_feature_has_an_example_with_two_steps_an_C0F08DC30BAD()
    'Given the feature has an example with two steps and a comment line between the steps
    TStepVars.feature_text = TStepVars.feature_text & vbLf _
                        & "  Example: sample example" & vbLf _
                        & "   Given step one" & vbLf _
                        & " # this is a comment" & vbLf _
                        & "   And step two"
End Sub

Public Sub Given_the_feature_has_a_background_clause_containi_EE3E39AB4582()
    'Given the feature has a background clause containing a Given step
        TStepVars.feature_text = TStepVars.feature_text & vbLf _
                        & "  Background:" & vbLf _
                        & "   Given a step"
End Sub
Public Sub When_the_feature_is_parsed_1AA9891E2CA3()
    'When the feature is parsed
    Set m_parsed_feature = TFeatureParser.parse_feature(TStepVars.feature_text)
End Sub

Public Sub Then_the_parsing_results_in_the_error_message_Feat_C2170B1F73BD()
    'Then the parsing results in the error message "Feature lacks feature keyword at the beginning"
    TSpec.expect(m_parsed_feature.ErrorStatus).to_be "Feature lacks feature keyword at the beginning"
End Sub

Public Sub Then_the_parsed_result_contains_a_separate_item_fo_6E788222165F()
    'Then the parsed result contains a separate item for each of the given elements
    TSpec.expect(m_parsed_feature.Name).to_be "sample feature"
    TSpec.expect(m_parsed_feature.Clauses.Count).to_be 2
    TSpec.expect(TypeName(m_parsed_feature.Clauses(1))).to_be "TRule"
    TSpec.expect(m_parsed_feature.Clauses(1).Name).to_be "this is a sample rule"
    TSpec.expect(TypeName(m_parsed_feature.Clauses(2))).to_be "TExample"
    TSpec.expect(m_parsed_feature.Clauses(2).Name).to_be "this is an example"
End Sub

Public Sub Then_the_parsed_result_contains_a_feature_with_the_793176670299()
    'Then the parsed result contains a feature with the name "sample feature"
    TSpec.expect(m_parsed_feature.Name).to_be "sample feature"
End Sub

Public Sub Then_the__parsed_result_contains_a_rule_with_the_n_7F67956919B9()
    'Then the  parsed result contains a rule with the name "this is a rule"
    TSpec.expect(TypeName(m_parsed_feature.Clauses(1))).to_be "TRule"
    TSpec.expect(m_parsed_feature.Clauses(1).Name).to_be "this is a rule"
End Sub

Public Sub Then_the_feature_description_is_set_to_those_two_l_A335E10C9C02()
    'Then the feature description is set to those two lines
    TSpec.expect(m_parsed_feature.description).to_be "this is" & vbLf & "the description"
End Sub

Public Sub Then_the_parsed_feature_contains_a_rule_5FF16597B7E1()
    'Then the parsed feature contains a rule
    TSpec.expect(TypeName(m_parsed_feature.Clauses(1))).to_be "TRule"
End Sub

Public Sub Then_the_rules_description_is_set_to_this_is_a_des_B3D47E4F8255()
    'Then the rules description is set to "this is a description"
     TSpec.expect(m_parsed_feature.Clauses(1).description).to_be "this is a description"
End Sub

Public Sub Then_the_parsed_features_contains_an_example_AF5E191DE8EE()
    'Then the parsed features contains an example
    TSpec.expect(TypeName(m_parsed_feature.Clauses(1))).to_be "TExample"
    Set TStepVars.example = m_parsed_feature.Clauses(1)
End Sub

Public Sub Then_the_example_clause_from_the_parsed_result_con_F111F3E2BACC()
    'Then the example clause from the parsed result contains all the steps
    
    Dim example_steps As Collection
    
    Set example_steps = m_parsed_feature.Clauses(1).Steps
    TSpec.expect(example_steps.Count).to_be 3
    TSpec.expect(example_steps(1).SType).to_be STEP_TYPE_GIVEN
    TSpec.expect(example_steps(1).Name).to_be "a precondition"
    TSpec.expect(example_steps(1).Head).to_be "Given"
    TSpec.expect(example_steps(2).SType).to_be STEP_TYPE_WHEN
    TSpec.expect(example_steps(2).Name).to_be "action happens"
    TSpec.expect(example_steps(2).Head).to_be "When"
    TSpec.expect(example_steps(3).SType).to_be STEP_TYPE_THEN
    TSpec.expect(example_steps(3).Name).to_be "some result is expected"
    TSpec.expect(example_steps(3).Head).to_be "Then"
End Sub

Public Sub Then_the_parsed_feature_doesnt_contain_any_error_0453F9190B1C()
    'Then the parsed feature doesn't contain any error
    TSpec.expect(m_parsed_feature.ErrorStatus).to_be vbNullString
End Sub

Public Sub Then_the_parsed_features_contains_the_tags_wip_imp_3E416E210DFE()
    'Then the parsed features contains the tags wip, important and beta
    TSpec.expect(m_parsed_feature.Tags("@wip")).to_be "@wip"
    TSpec.expect(m_parsed_feature.Tags("@important")).to_be "@important"
    TSpec.expect(m_parsed_feature.Tags("@beta")).to_be "@beta"
End Sub

Public Sub Then_the_example_contains_the_tags_wip_important_a_C2643DFE6EC5()
    'Then the example contains the tags wip, important and beta
    TSpec.expect(TStepVars.example.Tags("@wip")).to_be "@wip"
    TSpec.expect(TStepVars.example.Tags("@important")).to_be "@important"
    TSpec.expect(TStepVars.example.Tags("@beta")).to_be "@beta"
End Sub

Public Sub Then_the_parsed_result_contains_an_example_with_tw_7F3806C826C5()
    'Then the parsed result contains an example with two steps
    TSpec.expect(m_parsed_feature.Clauses(1).Steps.Count).to_be 2
End Sub

Public Sub Then_the_Given_step_is_assigned_to_the_feature_622D49FE1220()
    'Then the Given step is assigned to the feature
    TSpec.expect(m_parsed_feature.BackgroundSteps(1).StepDefinition).to_be "Given a step"
End Sub

