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

Public Sub Given_an_example_STR_6A35400021EE(example_context As TContext, step_expressions As Collection)
    'Given an example {string}
    example_context.set_value Trim(step_expressions(1)), "gherkin_text"
End Sub

Public Sub Given_this_line_is_followed_by_STR_A388704A4A7E(example_context As TContext, step_expressions As Collection)
    'Given this line is followed by "    Given a precondition"
    example_context.set_value example_context.get_value("gherkin_text") & vbLf & step_expressions(1), "gherkin_text"
End Sub

Public Sub Given_an_example_with_two_steps_STR_and_STR_A2651FF47542(example_context As TContext, step_expressions As Collection)
    'Given an example with two steps {string} and {string}
    example_context.set_value "Feature: sample feature" & vbLf _
                        & vbLf _
                        & "  Example: sample example" & vbLf _
                        & step_expressions(1) & vbLf _
                        & step_expressions(2), "gherkin_text"
End Sub

Public Sub Given_the_first_step_is_STR_5F91EBC94964(example_context As TContext, step_expressions As Collection)
    'Given the first step is "Given a first step"
    example_context.set_value example_context.get_value("gherkin_text") & vbLf _
                                & step_expressions(1), "gherkin_text"
End Sub

Public Sub Given_a_background_6A351EBB4211(example_context As TContext)
    'Given a background
    example_context.set_value "Background:", "gherkin_text"
End Sub

Public Sub Then_the_parsed_result_contains_an_example_7B67C585A7A0(example_context As TContext)
    'Then the parsed result contains an example
    TSpec.expect(TypeName(example_context.get_value("parsed_feature").sections(1))).to_be "TExample"
End Sub

Public Sub When_the_example_is_parsed_CEB09D77F01F(example_context As TContext)
    'When the example is parsed
    Dim Example As TExample
     
    On Error GoTo exception_handler
    Set Example = TFeatureParser.parse_example(example_context.get_value("gherkin_text"), THelper.new_TFeature, Nothing)
    example_context.set_value Example, "parsed_example"
    Exit Sub
    
exception_handler:
    example_context.set_value Err.description, "err_msg"
End Sub

Public Sub Then_the_parsed_example_contains_all_the_steps_013B99DC739C(example_context As TContext)
    'Then the parsed example contains all the steps
    Dim Example As TExample
    
    Set Example = example_context.get_value("parsed_example")
    TSpec.expect(Example.steps.Count).to_be 3
    TSpec.expect(Example.steps(1).SType).to_be STEP_TYPE_GIVEN
    TSpec.expect(Example.steps(1).Name).to_be "a precondition"
    TSpec.expect(Example.steps(1).Elements(1)).to_be "Given"
    TSpec.expect(Example.steps(2).SType).to_be STEP_TYPE_WHEN
    TSpec.expect(Example.steps(2).Name).to_be "action happens"
    TSpec.expect(Example.steps(2).Elements(1)).to_be "When"
    TSpec.expect(Example.steps(3).SType).to_be STEP_TYPE_THEN
    TSpec.expect(Example.steps(3).Name).to_be "some result is expected"
    TSpec.expect(Example.steps(3).Elements(1)).to_be "Then"
End Sub

Public Sub Then_the_type_of_the_second_step_is_set_to_STR_9CA91DFBD6DD(example_context As TContext, step_expressions As Collection)
    'Then the type of the second step is set to {string}
    TSpec.expect(example_context.get_value("parsed_feature").sections(1).steps(2).SType).to_be step_expressions(1)
End Sub

Public Sub Then_first_step_is_changed_to_STR_26F73CDF60C2(example_context As TContext, step_expressions As Collection)
    'Then first step is changed to {string}
    TSpec.expect(example_context.get_value("example").steps(1).StepDefinition).to_be step_expressions(1)
End Sub

