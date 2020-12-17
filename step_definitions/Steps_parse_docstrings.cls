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

Public Sub Given_an_example_definition_2803C497D1D1()
    'Given an example definition
    TStepVars.gherkin_text = "Example: docstring sample"
End Sub

Public Sub Given_this_step_is_followed_by_a_docstring_contain_C4061F67175D(step_expressions As Collection)
    'Given this step is followed by a docstring containing "this is a docstring"
    TStepVars.gherkin_text = TStepVars.gherkin_text & vbLf _
                                & """""""" & vbLf _
                                & step_expressions(1) & vbLf _
                                & """"""""
End Sub

Public Sub Given_this_is_followed_by_a_docstring_STR_CC499A469EEE(step_expressions As Collection)
    'And this is followed by a docstring {string}
    TStepVars.gherkin_text = TStepVars.gherkin_text & vbLf _
                                & """""""" & vbLf _
                                & step_expressions(1) & vbLf _
                                & """"""""
End Sub

Public Sub Given_an_example_definition_STR_2AA8F75D1F34(step_expressions As Collection)
    'Given an example definition {string}
    TStepVars.gherkin_text = step_expressions(1)
End Sub

Public Sub Then_the_first_step_has_an_expression_451F9158BDF6()
    'Then the first step has an expression
    TSpec.expect(TStepVars.first_step.Expressions.Count).to_be 1
End Sub

Public Sub Then_the_function_name_for_the_first_step_ends_wit_79EB11152D22(step_expressions As Collection)
    'And the function name for the first step ends with {string}
    Dim function_name As String
    
    function_name = TStepVars.example.Steps(1).get_step_function_name
    function_name = Left(function_name, Len(function_name) - 13)
    TSpec.expect(Right(function_name, 3)).to_be step_expressions(1)
End Sub

Public Sub Then_the_step_has_one_string_parameter_with_two_li_AD35FD5BD215()
    'Then the step has one string parameter with two lines
    TSpec.expect(TypeName(TStepVars.example.Steps(1).Expressions(1))).to_be "String"
    TSpec.expect(UBound(Split(TStepVars.example.Steps(1).Expressions(1), vbLf))).to_be 1
End Sub
