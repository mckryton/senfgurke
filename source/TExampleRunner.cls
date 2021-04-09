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

Public Sub run_example(example As TExample)
    
    Dim step_result As Variant
    Dim err_msg As String
    Dim step As TStep
    Dim step_status As String
    Dim previous_step_status As String
    Dim step_runner As TStepRunner

    step_result = vbNullString
    Set step_runner = New TStepRunner
    step_runner.setup m_session
    step_status = vbNullString
    m_session.Reporter.report REPORT_MSG_TYPE_EXAMPLE_TITLE, example.OriginalHeadline
    For Each step In example.Steps
        err_msg = vbNullString
        previous_step_status = step_status
        step_result = step_runner.run_step(step)
        step_status = CStr(step_result(0))
        If previous_step_status <> STATUS_OK And previous_step_status <> vbNullString Then
            ' mark all steps as skipped when the last step wasn't OK
            If step_status <> STATUS_MISSING Then step_status = STATUS_SKIPPED
        End If
        If Not step_status = STATUS_OK And Not step_status = STATUS_SKIPPED Then
            'step failed -> get the failure message
            err_msg = step_result(1)
            'this is a hack because of the broken exception handling in vba (e.g. vba will overwrite custom err descriptions)
            If step_status = STATUS_PENDING Then err_msg = TSpec.LastFailMsg
        End If
        m_session.statistics.count_event LOG_TYPE_STEP, step_status, step.OriginalStepDefinition, ExtraVBA.get_unix_timestamp_now
        m_session.Reporter.report REPORT_MSG_TYPE_STEP, step.OriginalStepDefinition, step_status, err_msg
    Next
    'the last step status sets the status for the example
    m_session.statistics.count_event LOG_TYPE_EXAMPLE, step_status, example.OriginalHeadline, ExtraVBA.get_unix_timestamp_now
End Sub

