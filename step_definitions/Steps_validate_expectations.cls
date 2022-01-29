VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_validate_expectations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_a_variable_refers_to_an_object_C82E72756271(example_context As TContext)
    'Given a variable refers to an object
    example_context.set_value New Collection, "sample_variable"
End Sub

Public Sub Given_a_variable_refers_to_nothing_6F7619677868(example_context As TContext)
    'Given a variable refers to Nothing
    example_context.set_value Nothing, "sample_variable"
End Sub

Public Sub Given_a_variable_refers_to_INT_B6E787C91FDE(example_context As TContext, step_expressions As Collection)
    'Given a variable refers to {integer}
    example_context.set_value step_expressions(1), "sample_variable"
End Sub

Public Sub When_the_variable_is_expected_to_be_nothing_DBE6C575C088(example_context As TContext)
    'When the variable is expected to be Nothing
    On Error GoTo expectation_failed
    TSpec.expect(example_context.get_value("sample_variable")).to_be_nothing
    example_context.set_value "confirmed", "expectation_result"
    Exit Sub
    
expectation_failed:
    example_context.set_value "failed", "expectation_result"
End Sub

Public Sub When_the_nonobject_variable_is_expected_to_be_noth_10954F3DA4DC(example_context As TContext)
    'When the non-object variable is expected to be Nothing
    On Error GoTo expectation_failed
    TSpec.expect(example_context.get_value("sample_variable")).to_be_nothing
    example_context.set_value "confirmed", "expectation_result"
    Exit Sub
    
expectation_failed:
    example_context.set_value "failed", "expectation_result"
End Sub

Public Sub When_the_variable_is_expected_not_to_be_nothing_63DB3DDBA289(example_context As TContext)
    'When the variable is expected not to be Nothing
    On Error GoTo expectation_failed
    TSpec.expect(example_context.get_value("sample_variable")).not_to_be_nothing
    example_context.set_value "confirmed", "expectation_result"
    Exit Sub
    
expectation_failed:
    example_context.set_value "failed", "expectation_result"
End Sub

Public Sub When_the_nonobject_variable_is_expected_not_to_be__A409ACACC890(example_context As TContext)
    'When the non-object variable is expected not to be Nothing
    On Error GoTo expectation_failed
    TSpec.expect(example_context.get_value("sample_variable")).to_be_nothing
    example_context.set_value "confirmed", "expectation_result"
    Exit Sub
    
expectation_failed:
    example_context.set_value "failed", "expectation_result"
End Sub

Public Sub Then_the_expectation_is_confirmed_A341C2A69D10(example_context As TContext)
    'Then the expectation is confirmed
    Dim expectation_result As String
    
    expectation_result = example_context.get_value("expectation_result")
    TSpec.expect(expectation_result).to_be "confirmed"
End Sub

Public Sub Then_the_expectation_fails_DBF0F873A79F(example_context As TContext)
    'Then the expectation fails
    Dim expectation_result As String
    
    expectation_result = example_context.get_value("expectation_result")
    TSpec.expect(expectation_result).to_be "failed"
End Sub

