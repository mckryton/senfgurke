VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_parse_step_expressions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Then_the_step_has_one_step_expression_with_the_dat_90F0EF0E4524(step_expressions As Collection)
    'Then the step has one step expression with the data type {string}
    TSpec.expect(LCase(TypeName(TStepVars.step.Expressions(1)))).to_be step_expressions(1)
End Sub

Public Sub Then_the_value_of_the_expression_is_STR_8D490483B82D(step_expressions As Collection)
    'Then the value of the expression is {string}
    TSpec.expect(TStepVars.step.Expressions(1)).to_be step_expressions(1)
End Sub

Public Sub Then_the_value_of_the_expression_is_INT_8D4904831D87(step_expressions As Collection)
    'Then the value of the expression is {integer}
    TSpec.expect(TStepVars.step.Expressions(1)).to_be step_expressions(1)
End Sub

Public Sub Then_the_value_of_the_expression_is_DBL_8D490483D413(step_expressions As Collection)
    'Then the value of the expression is {double}
    TSpec.expect(TStepVars.step.Expressions(1)).to_be step_expressions(1)
End Sub

Public Sub Then_the_step_has_no_step_expressions_DF47F903F69C()
    'Then the step has no step expressions
    TSpec.expect(TStepVars.step.Expressions.Count).to_be 0
End Sub

Public Sub Then_the_step_has_INT_step_expressions_DF47F8B87951(step_expressions As Collection)
    'Then the step has {integer} step expressions
    TSpec.expect(TStepVars.step.Expressions.Count).to_be step_expressions(1)
End Sub

Public Sub Then_the_data_types_of_the_expressions_are_STR_F48BBED5B618(step_expressions As Collection)
    'And the data types of the expressions are {string}
    
    Dim data_types As Variant
    Dim data_type As Variant
    Dim index As Integer
    
    index = 1
    data_types = Split(step_expressions(1), ",")
    For Each data_type In data_types
        TSpec.expect(LCase(TypeName(TStepVars.step.Expressions(index)))).to_be LCase(Trim(data_type))
        index = index + 1
    Next
End Sub

Public Sub Then_the_value_of_the_expressions_are_STR_8D4904837C42(step_expressions As Collection)
    'And the value of the expressions are {string}
        
    Dim expr_values As Variant
    Dim expr_value As Variant
    Dim index As Integer
    
    index = 1
    expr_values = Split(step_expressions(1), ",")
    For Each expr_value In expr_values
        TSpec.expect(CStr(TStepVars.step.Expressions(index))).to_be Trim(expr_value)
        index = index + 1
    Next
End Sub
