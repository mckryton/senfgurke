VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TExampleRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_is_test_stopped As Boolean

Public Function run_example(example As TExample, Optional silent) As Collection

    Dim step_result As Variant
    Dim err_msg As String
    Dim step As TStep
    Dim example_statistics As Collection

    If IsMissing(silent) Then silent = False
    Set example_statistics = New Collection
    IsTestStopped = False
    If Not silent Then TRun.Reporter.report REPORT_MSG_TYPE_EXAMPLE_TITLE, example.OriginalHeadline
    example_statistics.Add example.OriginalHeadline, "example"
    On Error GoTo error_handler
    For Each step In example.Steps
        err_msg = vbNullString
        step_result = TStepRunner.run_step(step)
        If Not step_result(0) = STATUS_OK Then
            err_msg = step_result(1)
            'this is a hack because of the broken exception handling in vba (e.g. vba will overwrite custom err descriptions)
            If step_result(0) = STATUS_PENDING Then err_msg = TSpec.LastFailMsg
            stop_example
        End If
        example_statistics.Add Array(step_result(0), step.OriginalStepDefinition)
        If Not silent Then TRun.Reporter.report REPORT_MSG_TYPE_STEP, step.OriginalStepDefinition, step_result(0), err_msg
        If IsTestStopped Then
            Exit For
        End If
    Next
    Set run_example = example_statistics
    Exit Function
    
error_handler:
    TRun.Log.log_function_error "Runtime errror in TExampleRunner.run_example", _
        example(CLAUSE_ATTR_TYPE) & " " & example(CLAUSE_ATTR_NAME)
End Function

Public Sub stop_example()
    IsTestStopped = True
End Sub

Public Property Get IsTestStopped() As Boolean
    IsTestStopped = m_is_test_stopped
End Property

Private Property Let IsTestStopped(ByVal is_test_stopped As Boolean)
    m_is_test_stopped = is_test_stopped
End Property



