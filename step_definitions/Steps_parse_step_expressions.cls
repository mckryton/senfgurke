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

Public Sub Given_a_step_definition_STR_65172AC6DD8A(step_expressions As Collection)
    'Given a step definition {string}
    TStepVars.gherkin_text = step_expressions(1)
End Sub

Public Sub Then_the_step_has_one_step_expression_type_STR_E0D1BCAD3F0E(step_expressions As Collection)
    'Then the step has one step expression type {string}
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
