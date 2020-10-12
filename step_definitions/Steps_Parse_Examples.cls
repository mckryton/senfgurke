VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Parse_Examples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_a_feature_F5C67A506977()
    'Given a feature
    m_feature_text = "Feature: sample feature"
End Sub

Public Sub Given_the_feature_includes_a_line___Example_sample_AB5AFF90706C()
    'Given the feature includes a line "  Example: sample example"
    m_feature_text = m_feature_text & vbLf & "  Example: sample example"
End Sub

Public Sub Given_this_line_is_followed_by_____Given_a_precond_888463027745()
    'Given this line is followed by "    Given a precondition"
    m_feature_text = m_feature_text & vbLf & "    Given a precondition"
End Sub

Public Sub Given_this_line_is_followed_by_____When_action_hap_85C0334B9847()
    'Given this line is followed by "    When action happens"
    m_feature_text = m_feature_text & vbLf & "    When action happens"
End Sub

Public Sub Given_this_line_is_followed_by_____Then_some_resul_EA08DF64E8E9()
    'Given this line is followed by "    Then some result is expected"
    m_feature_text = m_feature_text & vbLf & "    Then some result is expected"
End Sub

Public Sub Given_an_example_with_two_steps_Given_x_is_1_and_A_192EBBAC25E6()
    'Given an example with two steps "Given x is 1" and "And y is 2"
    m_feature_text = "Feature: sample feature" & vbLf _
                        & vbLf _
                        & "  Example: sample example" & vbLf _
                        & "    Given x is 1" & vbLf _
                        & "    And y is 2"
End Sub

Public Sub Given_an_example_with_two_steps_When_some_action_h_BF2A304D0F15()
    'Given an example with two steps "When some action happens" and "But it doesn't matter"
    m_feature_text = "Feature: sample feature" & vbLf _
                    & vbLf _
                    & "  Example: sample example" & vbLf _
                    & "    When some action happens" & vbLf _
                    & "    But it doesn't matter"
End Sub

Public Sub Then_the_parsed_result_contains_an_example_AF5E3525167D()
    'Then the parsed result contains an example
    TSpec.expect(TypeName(m_parsed_feature.Clauses(1))).to_be "TExample"
End Sub

Public Sub Then_the_type_of_the_second_step_is_set_to_Given_83FF038C4F7A()
    'Then the type of the second step is set to "Given"
    TSpec.expect(m_parsed_feature.Clauses(1).Steps(2).SType).to_be STEP_TYPE_GIVEN
End Sub

Public Sub Then_the_type_of_the_second_step_is_set_to_When_83FF038C05FE()
    'Then the type of the second step is set to "When"
    TSpec.expect(m_parsed_feature.Clauses(1).Steps(2).SType).to_be STEP_TYPE_WHEN
End Sub
