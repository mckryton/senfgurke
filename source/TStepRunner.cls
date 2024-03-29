VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TStepRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_session As TSession
Private m_parent_runner As TExampleRunner

Private Sub Class_Terminate()
    Set m_session = Nothing
    Set m_parent_runner = Nothing
End Sub

Public Sub setup(session As TSession, parent_runner As TExampleRunner)
    Set m_session = session
    Set m_parent_runner = parent_runner
End Sub

Public Function run_step(step_definition As TStep, context As TContext, Optional dry_run) As Variant
    'run_step returns an array with status and optional error message
    Dim step_result As Variant
    Dim step_implementation_class As Variant
    Dim step_function_name As String
    Dim step_run_attempts As Integer
    
    If IsMissing(dry_run) Then dry_run = False
    If dry_run Then
        run_step = Array(STATUS_OK)
        Exit Function
    End If
    On Error GoTo error_handler
    step_run_attempts = 0
    step_function_name = step_definition.get_step_function_name
    For Each step_implementation_class In m_session.StepImplementations
        TSpec.LastFailMsg = vbNullString
        step_result = try_step(step_implementation_class, step_definition, context)
    Next
    If step_run_attempts = m_session.StepImplementations.Count Then
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
        run_step = fail_step(Err.Number, Err.description)
    End If
End Function

Public Function try_step(step_implementation_class As Variant, step_definition As TStep, context As TContext) As Variant
    Dim step_result As String
    Dim step_expressions As Collection
    Dim step_function_name As String
    
    step_function_name = step_definition.get_step_function_name
    Set step_expressions = ExtraVBA.copy_collection(step_definition.Expressions)
    If step_definition.docstring <> vbNullString Then step_expressions.Add step_definition.docstring
    If step_expressions.Count > 0 Then
        If step_definition.data_table Is Nothing Then
            step_result = CallByName(step_implementation_class, step_function_name, VbMethod, context, step_expressions)
        Else
            step_result = CallByName(step_implementation_class, step_function_name, VbMethod, context, step_expressions, step_definition.data_table)
        End If
    Else
        If step_definition.data_table Is Nothing Then
            step_result = CallByName(step_implementation_class, step_function_name, VbMethod, context)
        Else
            step_result = CallByName(step_implementation_class, step_function_name, VbMethod, context, step_definition.data_table)
        End If
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
        fail_step = Array(STATUS_MISSING, vbNullString)
        m_parent_runner.log_event EVENT_STEP_IS_MISSING, err_desc
    Case Else
        fail_step = Array(STATUS_FAIL, err_desc & vbLf & TSpec.LastFailMsg)
    End Select
End Function
