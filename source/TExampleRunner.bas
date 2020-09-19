Attribute VB_Name = "TExampleRunner"
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

Public Const EXAMPLE_ATTR_STEPS = "steps"

Dim m_is_test_stopped As Boolean

Public Function execute_step(step_definition As TStep, step_implementation As Variant) As Variant

    'TODO: refactor step implementation -> call steps indepent from theire class or module
    execute_step = step_implementation.run_step(step_definition.StepImplementation)
End Function

Public Function execute_example(example As Collection, feature_object As Variant) As Collection

    Dim step_index As Integer
    Dim step_result As Variant
    Dim err_msg As String
    Dim step As TStep
    Dim example_statistics As Collection

    On Error GoTo error_handler
    Set example_statistics = create_example_statistics
    TExampleRunner.IsTestStopped = False
    'step_index = 0
    feature_object.before_example
    'TODO: save original text from example after parsing
    TReport.report TReport.TYPE_EXAMPLE_TITLE, example(TFeatureRunner.CLAUSE_ATTR_TYPE) & ": " & example(TFeatureRunner.CLAUSE_ATTR_NAME)
    For Each step In example(TExampleRunner.EXAMPLE_ATTR_STEPS)
        err_msg = vbNullString
        step_result = execute_step(step, feature_object)
        If Not step_result(0) = "OK" Then
            err_msg = step_result(1)
            TExampleRunner.stop_example
        End If
        increase_stats_counter CStr(step_result(0)), example_statistics
        TReport.report TReport.TYPE_STEP, step(TExampleRunner.STEP_ATTR_HEAD) & " " & step(TExampleRunner.STEP_ATTR_NAME), step_result(0), err_msg
        If IsTestStopped Then
            Exit For
        End If
    Next
    If Not TExampleRunner_Old.IsTestStopped Then
        feature_object.after_example
    End If
    Set execute_example = example_statistics
    Exit Function
    
error_handler:
    If Err.Number = ERR_ID_SCENARIO_SYNTAX_ERROR Then
        TReport.Log.error_log "syntax error: " & Err.Description & vbCr & vbLf & "in line >" & _
            step(TExampleRunner.STEP_ATTR_HEAD) & " " & step(TExampleRunner.STEP_ATTR_NAME) & "<"
    Else
        TReport.Log.log_function_error "Runtime errror in TExampleRunner.execute_example", _
            example(TFeatureRunner.CLAUSE_ATTR_TYPE) & " " & example(TFeatureRunner.CLAUSE_ATTR_NAME)
    End If
End Function

Private Function create_example_statistics() As Collection
    
    Dim new_example_statistics As Collection
    
    Set new_example_statistics = New Collection
    new_example_statistics.Add 0, TExampleRunner.STEP_RESULT_OK
    new_example_statistics.Add 0, TExampleRunner.STEP_RESULT_MISSING
    new_example_statistics.Add 0, TExampleRunner.STEP_RESULT_PENDING
    new_example_statistics.Add 0, TExampleRunner.STEP_RESULT_FAIL
    Set create_example_statistics = new_example_statistics
End Function

Private Sub increase_stats_counter(result_type As String, example_statistics As Collection)

    Dim counter As Long
    
    'My kingdom for a hash!
    counter = example_statistics(result_type) + 1
    example_statistics.Remove result_type
    example_statistics.Add counter, result_type
End Sub
Public Sub stop_example()
    TExampleRunner.IsTestStopped = True
End Sub

Public Property Get IsTestStopped() As Boolean
    IsTestStopped = m_is_test_stopped
End Property

Private Property Let IsTestStopped(ByVal is_test_stopped As Boolean)
    m_is_test_stopped = is_test_stopped
End Property

Public Function fail_step(err_id As Long, Optional err_msg) As Variant
 
    Dim err_desc As String
    
    If IsMissing(err_msg) Then
        err_desc = vbNullString
    Else
        err_desc = err_msg
    End If
    Select Case err_id
    Case ERR_ID_STEP_IS_PENDING
        fail_step = Array(TExampleRunner.STEP_RESULT_PENDING, err_desc)
    Case ERR_ID_STEP_IS_MISSING
        fail_step = Array(TExampleRunner.STEP_RESULT_MISSING, "PENDING: add code snippet as suggestion for missing test steps")
    Case Else
        fail_step = Array(TExampleRunner.STEP_RESULT_FAIL, err_desc)
    End Select
End Function
