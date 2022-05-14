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

Public Sub Then_the_first_step_has_an_docstring_451F91589372(example_context As TContext)
    'Then the first step has an docstring
    Dim first_step As TStep
    
    Set first_step = example_context.get_value("step_list")(1)
    TSpec.expect(first_step.docstring).not_to_be vbNullString
End Sub

Public Sub Then_the_function_name_for_the_first_step_ends_wit_79EB11152D22(example_context As TContext, step_expressions As Collection)
    'And the function name for the first step ends with {string}
    Dim function_name As String
    Dim first_step As TStep
    
    Set first_step = example_context.get_value("step_list")(1)
    function_name = first_step.get_step_function_name
    function_name = Left(function_name, Len(function_name) - 13)
    TSpec.expect(Right(function_name, 3)).to_be step_expressions(1)
End Sub

Public Sub Then_the_docstring_of_the_first_step_is_STR_99C591958B2D(example_context As TContext, step_expressions As Collection)
    'And the docstring of the first step is
    Dim first_step As TStep
    
    Set first_step = example_context.get_value("step_list")(1)
    TSpec.expect(first_step.docstring).to_be step_expressions(1)
End Sub

Public Sub Then_the_first_line_of_the_steps_docstring_is_empt_C94363C3329C(example_context As TContext)
    'Then the first line of the steps docstring is empty
    Dim first_step As TStep
    
    Set first_step = example_context.get_value("step_list")(1)
    TSpec.expect(Split(first_step.docstring, vbLf)(0)).to_be vbNullString
End Sub
