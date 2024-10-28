VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_custom_err_msg_expectatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_an_expected_value_is_defined_as_INT_72FB69157348(example_context As TContext, step_expressions As Collection)
    'Given an expected value is defined as {integer}
    example_context.set_value step_expressions(1), "expected_val"
End Sub

Public Sub Given_the_actual_value_is_INT_5F9169C377F7(example_context As TContext, step_expressions As Collection)
    'And the actual value is {integer}
    example_context.set_value step_expressions(1), "actual_val"
End Sub

Public Sub Given_the_custom_error_message_is_defined_as__B811B39A8C79(example_context As TContext, step_expressions As Collection)
    'And the custom error message is defined as {string}
    example_context.set_value step_expressions(1), "custom_err_msg"
End Sub

Public Sub When_exptected_and_actual_value_are_being_com_CB74A1DA7D42(example_context As TContext)
    'When exptected and actual value are being compared using to_be
    Dim actual_err_msg As String
    Dim expected_val As Variant
    Dim actual_val As Variant
    
    actual_err_msg = vbNullString
    expected_val = example_context.get_value("expected_val")
    actual_val = example_context.get_value("actual_val")
    On Error GoTo comparison_failed
    TSpec.expect(actual_val).to_be expected_val
    On Error GoTo 0
    example_context.set_value actual_err_msg, "actual_err_msg"
    Exit Sub
    
comparison_failed:
    'sometimes VBA will overwrite any custom set Err.description property, therefore the err msg is copied to TSpec.LastFailMsg
    actual_err_msg = TSpec.LastFailMsg
    'remove the line break formatting from the default err msg
    actual_err_msg = Trim(Replace(actual_err_msg, vbLf, vbNullString))
    Err.Clear
    Resume Next
End Sub

Public Sub When_exptected_and_actual_value_are_being_com_404E2A120FDB(example_context As TContext)
    'When exptected and actual value are being compared using not_to_be
    pending
End Sub

Public Sub Then_the_expectation_fails_with_STR_as_error__AD5840755870(example_context As TContext, step_expressions As Collection)
    'Then the expectation fails with {string} as error description
    Dim expected_err_msg As String
    Dim actual_err_msg As String
    
   expected_err_msg = step_expressions(1)
   actual_err_msg = example_context.get_value("actual_err_msg")
   TSpec.expect(actual_err_msg).to_be expected_err_msg
End Sub


