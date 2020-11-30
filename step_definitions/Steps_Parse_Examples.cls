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

Dim m_parsed_example As TExample
Dim m_parsed_background As TBackground

Public Sub Given_a_feature_F5C668746977()
    'Given a feature
    TStepVars.gherkin_text = "Feature: sample feature"
End Sub

Public Sub Given_the_feature_includes_a_line_STR_84B3E9B86DF3(step_expressions As Collection)
    'Given the feature includes a line {string}
    TStepVars.gherkin_text = TStepVars.gherkin_text & vbLf & step_expressions(1)
End Sub

Public Sub Given_this_line_is_followed_by_STR_A388704A4A7E(step_expressions As Collection)
    'Given this line is followed by "    Given a precondition"
    TStepVars.gherkin_text = TStepVars.gherkin_text & vbLf & step_expressions(1)
End Sub

Public Sub Given_an_example_with_two_steps_STR_and_STR_A2651FF47542(step_expressions As Collection)
    'Given an example with two steps {string} and {string}
    TStepVars.gherkin_text = "Feature: sample feature" & vbLf _
                        & vbLf _
                        & "  Example: sample example" & vbLf _
                        & step_expressions(1) & vbLf _
                        & step_expressions(2)
End Sub

Public Sub Given_an_example_F5C6F737F747()
    'Given an example
    TStepVars.gherkin_text = "Example: sample example"
End Sub

Public Sub Given_the_first_step_is_STR_5F91EBC94964(step_expressions As Collection)
    'Given the first step is "Given a first step"
    TStepVars.gherkin_text = TStepVars.gherkin_text & vbLf _
                                & step_expressions(1)
End Sub

Public Sub Given_this_step_is_followed_by_a_docstring_contain_C4061F67175D(step_expressions As Collection)
    'Given this step is followed by a docstring containing "this is a docstring"
    TStepVars.gherkin_text = TStepVars.gherkin_text & vbLf _
                                & """""""" & vbLf _
                                & step_expressions(1) & vbLf _
                                & """"""""
End Sub

Public Sub Given_a_background_6A351EBB4211()
    'Given a background
    TStepVars.gherkin_text = "Background:"
End Sub

Public Sub Given_the_first_step_of_the_background_is_STR_BB1462E43AFA(step_expressions As Collection)
    'Given the first step of the background is {string}
    TStepVars.gherkin_text = TStepVars.gherkin_text & vbLf & step_expressions(1)
End Sub

Public Sub Then_the_parsed_result_contains_an_example_7B67C585A7A0()
    'Then the parsed result contains an example
    TSpec.expect(TypeName(m_parsed_feature.Clauses(1))).to_be "TExample"
End Sub

Public Sub When_the_example_is_parsed_CEB09D77F01F()
    'When the example is parsed
    Set m_parsed_example = New TExample
    TFeatureParser.parse_steps Split(Trim(TStepVars.gherkin_text), vbLf), 0, m_parsed_example
End Sub

Public Sub When_the_feature_background_is_parsed_8C5FB8584F5D()
    'When the feature background is parsed
    Set m_parsed_background = New TBackground
    TFeatureParser.parse_steps Split(Trim(TStepVars.gherkin_text), vbLf), 0, m_parsed_background
End Sub

Public Sub Then_the_example_clause_from_the_parsed_result_con_C8ADD15947E9()
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

Public Sub Then_the_type_of_the_second_step_is_set_to_STR_9CA91DFBD6DD(step_expressions As Collection)
    'Then the type of the second step is set to {string}
    TSpec.expect(m_parsed_feature.Clauses(1).Steps(2).SType).to_be step_expressions(1)
End Sub

Public Sub Then_first_step_is_changed_to_STR_26F73CDF60C2(step_expressions As Collection)
    'Then first step is changed to {string}
    TSpec.expect(m_parsed_example.Steps(1).StepDefinition).to_be step_expressions(1)
End Sub

Public Sub Then_first_step_of_the_background_is_changed_to_ST_43C166035DB0(step_expressions As Collection)
    'Then first step of the background is changed to {string}
    TSpec.expect(m_parsed_background.Steps(1).StepDefinition).to_be step_expressions(1)
End Sub
