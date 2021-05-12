VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_parse_docstrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_an_example_definition_2803C497D1D1(example_context As TContext)
    'Given an example definition
    example_context.set_value "Example: docstring sample", "gherkin_text"
End Sub

Public Sub Given_the_step_is_followed_by_a_docstring_STR_FFEF53DE9EEE(example_context As TContext, step_expressions As Collection)
    'And the step is followed by a docstring {string}
    example_context.set_value example_context.get_value("gherkin_text") & vbLf _
                                & """""""" & vbLf _
                                & step_expressions(1) & vbLf _
                                & """""""", "gherkin_text"
End Sub

Public Sub Given_the_step_is_followed_by_a_docstring_with_a_l_716F6D071657(example_context As TContext)
    'And the step is followed by a docstring with a leading linebreak
    example_context.set_value example_context.get_value("gherkin_text") & vbLf _
                            & """""""" & vbLf _
                            & vbLf _
                            & "sample docstring" & vbLf _
                            & """""""", "gherkin_text"
End Sub

Public Sub Given_an_example_definition_STR_2AA8F75D1F34(example_context As TContext, step_expressions As Collection)
    'Given an example definition {string}
    example_context.set_value step_expressions(1), "gherkin_text"
End Sub

Public Sub When_the_step_is_parsed_317375497656(example_context As TContext)
    'When the step is parsed
    example_context.set_value New TExample, "example"
    example_context.set_value "Example: sample example" & vbLf & example_context.get_value("gherkin_text"), "gherkin_text"
    TFeatureParser.parse_steps Split(example_context.get_value("gherkin_text"), vbLf), 0, example_context.get_value("example")
End Sub

Public Sub When_the_feature_background_is_parsed_8C5FB8584F5D(example_context As TContext)
    'When the feature background is parsed
    Dim background As TBackground
    
    Set background = New TBackground
    TFeatureParser.parse_steps Split(Trim(example_context.get_value("gherkin_text")), vbLf), 0, background
    If background.Steps.Count = 1 Then
        example_context.set_value background.Steps(1), "first_step"
    End If
End Sub

Public Sub Then_the_first_step_has_an_expression_451F9158BDF6(example_context As TContext)
    'Then the first step has an expression
    TSpec.expect(example_context.get_value("first_step").Expressions.Count).to_be 1
End Sub

Public Sub Then_the_first_step_has_a_docstring_value_STR_B59597EF055B(example_context As TContext, step_expressions As Collection)
    'And the first step has a docstring value {string}
    TSpec.expect(example_context.get_value("first_step").Docstring).to_be step_expressions(1)
End Sub

Public Sub Then_the_function_name_for_the_first_step_ends_wit_79EB11152D22(example_context As TContext, step_expressions As Collection)
    'And the function name for the first step ends with {string}
    Dim function_name As String
    
    If Not example_context.value_exists("first_step") Then example_context.set_value example_context.get_value("example").Steps(1), "first_step"
    function_name = example_context.get_value("first_step").get_step_function_name
    function_name = Left(function_name, Len(function_name) - 13)
    TSpec.expect(Right(function_name, 3)).to_be step_expressions(1)
End Sub

Public Sub Then_the_step_has_one_string_parameter_with_two_li_AD35FD5BD215(example_context As TContext)
    'Then the step has one string parameter with two lines
    TSpec.expect(TypeName(example_context.get_value("example").Steps(1).Expressions(1))).to_be "String"
    TSpec.expect(UBound(Split(example_context.get_value("example").Steps(1).Expressions(1), vbLf))).to_be 1
End Sub

Public Sub Then_line_INT_of_the_steps_docstring_is_empty_FBE59787C9D9(example_context As TContext, step_expressions As Collection)
    'Then line {integer} of the steps docstring is empty
    TSpec.expect(Split(example_context.get_value("example").Steps(1).Docstring, vbLf)(step_expressions(1) - 1)).to_be vbNullString
End Sub
