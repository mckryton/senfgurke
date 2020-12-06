Attribute VB_Name = "TStepRunner"
Option Explicit

Public Function run_step(step_definition As TStep) As Variant

    Dim step_result As Variant
    Dim step_implementation_class As Variant
    Dim step_function_name As String
    Dim step_run_attempts As Integer
    Dim step_parameter_values As Variant
    
    On Error GoTo error_handler
    step_run_attempts = 0
    step_function_name = step_definition.get_step_function_name
    For Each step_implementation_class In TConfig.StepImplementations
        TSpec.LastFailMsg = vbNullString
        step_result = try_step(step_implementation_class, step_function_name, step_definition.Expressions)
    Next
    If step_run_attempts = TConfig.StepImplementations.Count Then
        run_step = fail_step(ERR_ID_STEP_IS_STATUS_MISSING, step_definition.get_step_function_template)
    Else
        run_step = Array(STATUS_OK)
    End If
    Exit Function

error_handler:
    If Err.Number = 438 Then
        step_run_attempts = step_run_attempts + 1
        Resume Next
    Else
        run_step = fail_step(Err.Number, Err.Description)
    End If
End Function

Private Function try_step(step_implementation_class As Variant, step_function_name As String, step_expressions As Collection) As Variant

    Dim step_result As String
    
    If step_expressions.Count > 0 Then
        step_result = CallByName(step_implementation_class, step_function_name, VbMethod, step_expressions)
    Else
        step_result = CallByName(step_implementation_class, step_function_name, VbMethod)
    End If
    try_step = step_result
End Function


Public Function fail_step(err_id As Long, Optional err_msg) As Variant
 
    Dim err_desc As String
    
    If IsMissing(err_msg) Then
        err_desc = vbNullString
    Else
        err_desc = err_msg
    End If
    Select Case err_id
    Case ERR_ID_STEP_IS_STATUS_PENDING
        fail_step = Array(STATUS_PENDING, err_desc)
    Case ERR_ID_STEP_IS_STATUS_MISSING
        fail_step = Array(STATUS_MISSING, err_desc)
    Case Else
        fail_step = Array(STATUS_FAIL, err_desc & vbLf & TSpec.LastFailMsg)
    End Select
End Function

Public Sub pending(Optional pending_msg)

    If IsMissing(pending_msg) Then pending_msg = vbNullString
    TSpec.LastFailMsg = pending_msg
    Err.Raise ERR_ID_STEP_IS_STATUS_PENDING, Description:=pending_msg
End Sub

