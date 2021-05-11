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
    TStepVars.gherkin_text = "Example: docstring sample"
End Sub

Public Sub Given_the_step_is_followed_by_a_docstring_STR_FFEF53DE9EEE(example_context As TContext, step_expressions As Collection)
    'And the step is followed by a docstring {string}
    TStepVars.gherkin_text = TStepVars.gherkin_text & vbLf _
                                & """""""" & vbLf _
                                & step_expressions(1) & vbLf _
                                & """"""""
End Sub

Public Sub Given_the_step_is_followed_by_a_docstring_with_a_l_716F6D071657(example_context As TContext)
    'And the step is followed by a docstring with a leading linebreak
        TStepVars.gherkin_text = TStepVars.gherkin_text & vbLf _
                                & """""""" & vbLf _
                                & vbLf _
                                & "sample docstring" & vbLf _
                                & """"""""
End Sub

Public Sub Given_an_example_definition_STR_2AA8F75D1F34(example_context As TContext, step_expressions As Collection)
    'Given an example definition {string}
    TStepVars.gherkin_text = step_expressions(1)
End Sub

Public Sub When_the_step_is_parsed_317375497656(example_context As TContext)
    'When the step is parsed
    Set TStepVars.example = New TExample
    TStepVars.gherkin_text = "Example: sample example" & vbLf & TStepVars.gherkin_text
    TFeatureParser.parse_steps Split(TStepVars.gherkin_text, vbLf), 0, TStepVars.example
End Sub

Public Sub When_the_feature_background_is_parsed_8C5FB8584F5D(example_context As TContext)
    'When the feature background is parsed
    Set TStepVars.background = New TBackground
    TFeatureParser.parse_steps Split(Trim(TStepVars.gherkin_text), vbLf), 0, TStepVars.background
    If TStepVars.background.Steps.Count = 1 Then
        Set TStepVars.first_step = TStepVars.background.Steps(1)
    End If
End Sub

Public Sub Then_the_first_step_has_an_expression_451F9158BDF6(example_context As TContext)
    'Then the first step has an expression
    TSpec.expect(TStepVars.first_step.Expressions.Count).to_be 1
End Sub

Public Sub Then_the_first_step_has_a_docstring_value_STR_B59597EF055B(example_context As TContext, step_expressions As Collection)
    'And the first step has a docstring value {string}
    TSpec.expect(TStepVars.first_step.Docstring).to_be step_expressions(1)
End Sub

Public Sub Then_the_function_name_for_the_first_step_ends_wit_79EB11152D22(example_context As TContext, step_expressions As Collection)
    'And the function name for the first step ends with {string}
    Dim function_name As String
    
    If TStepVars.first_step Is Nothing Then Set TStepVars.first_step = TStepVars.example.Steps(1)
    function_name = TStepVars.first_step.get_step_function_name
    function_name = Left(function_name, Len(function_name) - 13)
    TSpec.expect(Right(function_name, 3)).to_be step_expressions(1)
End Sub

Public Sub Then_the_step_has_one_string_parameter_with_two_li_AD35FD5BD215(example_context As TContext)
    'Then the step has one string parameter with two lines
    TSpec.expect(TypeName(TStepVars.example.Steps(1).Expressions(1))).to_be "String"
    TSpec.expect(UBound(Split(TStepVars.example.Steps(1).Expressions(1), vbLf))).to_be 1
End Sub

Public Sub Then_line_INT_of_the_steps_docstring_is_empty_FBE59787C9D9(example_context As TContext, step_expressions As Collection)
    'Then line {integer} of the steps docstring is empty
    TSpec.expect(Split(TStepVars.example.Steps(1).Docstring, vbLf)(step_expressions(1) - 1)).to_be vbNullString
End Sub
