Attribute VB_Name = "TStepRunner"
Option Explicit

Public Const STEP_TYPE_GIVEN = "Given"
Public Const STEP_TYPE_WHEN = "When"
Public Const STEP_TYPE_THEN = "Then"

Public Const STEP_ATTR_TYPE = "step_type"
Public Const STEP_ATTR_HEAD = "step_head"
Public Const STEP_ATTR_NAME = "step_name"
Public Const STEP_ATTR_ERROR = "error"

Public Const STEP_RESULT_OK = "OK"
Public Const STEP_RESULT_FAIL = "FAIL"
Public Const STEP_RESULT_MISSING = "MISSING"
Public Const STEP_RESULT_PENDING = "PENDING"

Public Function run_step(step_definition As TStep) As Variant ', step_implementation As Variant) As Variant

    Dim step_result As Variant
    Dim step_implementation_class As Variant
    Dim step_function_name As String
    Dim step_run_attempts As Integer
    
    On Error GoTo error_handler
    step_run_attempts = 0
    step_function_name = step_definition.get_step_function_name
    For Each step_implementation_class In TConfig.StepImplementations
        TSpec.LastFailMsg = vbNullString
        step_result = CallByName(step_implementation_class, step_function_name, VbMethod)
    Next
    If step_run_attempts = TConfig.StepImplementations.Count Then
        run_step = fail_step(ERR_ID_STEP_IS_MISSING, step_definition.get_step_function_template)
    Else
        run_step = Array("OK")
    End If
    Exit Function

error_handler:
    If Err.Number = 438 Then
        step_run_attempts = step_run_attempts + 1
        Resume Next
    Else
        run_step = fail_step(Err.Number, Err.description)
    End If
End Function

Public Function fail_step(err_id As Long, Optional err_msg) As Variant
 
    Dim err_desc As String
    
    If IsMissing(err_msg) Then
        err_desc = vbNullString
    Else
        err_desc = err_msg
    End If
    Select Case err_id
    Case ERR_ID_STEP_IS_PENDING
        fail_step = Array(STEP_RESULT_PENDING, err_desc)
    Case ERR_ID_STEP_IS_MISSING
        fail_step = Array(STEP_RESULT_MISSING, err_desc)
    Case Else
        fail_step = Array(STEP_RESULT_FAIL, err_desc & vbLf & TSpec.LastFailMsg)
    End Select
End Function

Public Sub pending(pending_msg)
    Err.Raise ERR_ID_STEP_IS_PENDING, description:=pending_msg
End Sub

