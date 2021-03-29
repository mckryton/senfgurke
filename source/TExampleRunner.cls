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

Private m_session As TSession

Public Sub setup(Session As TSession)
    Set m_session = Session
End Sub

Public Function run_example(example As TExample, Optional silent) As Collection

    Dim step_result As Variant
    Dim err_msg As String
    Dim step As TStep
    Dim example_statistics As Collection
    Dim last_step_result As String
    Dim step_runner As TStepRunner

    If IsMissing(silent) Then silent = False
    Set step_runner = New TStepRunner
    step_runner.setup m_session
    step_result = Array(vbNullString, vbNullString)
    Set example_statistics = New Collection
    m_session.Reporter.report REPORT_MSG_TYPE_EXAMPLE_TITLE, example.OriginalHeadline, silent:=silent
    example_statistics.Add example.OriginalHeadline, "example"
    For Each step In example.Steps
        err_msg = vbNullString
        last_step_result = step_result(0)
        step_result = step_runner.run_step(step, silent:=silent)
        ' mark all steps as skipped when the last step wasn't OK
        If last_step_result <> STATUS_OK And last_step_result <> vbNullString Then
            If step_result(0) <> STATUS_MISSING Then step_result(0) = STATUS_SKIPPED
        End If
        If Not step_result(0) = STATUS_OK And Not step_result(0) = STATUS_SKIPPED Then
            err_msg = step_result(1)
            'this is a hack because of the broken exception handling in vba (e.g. vba will overwrite custom err descriptions)
            If step_result(0) = STATUS_PENDING Then err_msg = TSpec.LastFailMsg
        End If
        example_statistics.Add Array(step_result(0), step.OriginalStepDefinition)
        m_session.Reporter.report REPORT_MSG_TYPE_STEP, step.OriginalStepDefinition, step_result(0), err_msg, silent
    Next
    Set run_example = example_statistics
End Function

