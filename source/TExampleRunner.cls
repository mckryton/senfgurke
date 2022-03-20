VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TExampleRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_session As TSession
Private m_parent_runner As TFeatureRunner

Private Sub Class_Terminate()
    Set m_session = Nothing
    Set m_parent_runner = Nothing
End Sub

Public Sub setup(session As TSession, parent_runner As TFeatureRunner)
    Set m_session = session
    Set m_parent_runner = parent_runner
End Sub

Public Sub run_example(example As TExample)
    Dim step_result As Variant
    Dim err_msg As String
    Dim step As TStep
    Dim step_status As String
    Dim previous_step_status As String
    Dim step_runner As TStepRunner
    Dim context As TContext

    step_result = vbNullString
    Set step_runner = New TStepRunner
    step_runner.setup m_session, Me
    Set context = New TContext
    step_status = vbNullString
    log_event EVENT_RUN_EXAMPLE_STARTED, example.description, example.OriginalHeadline, SECTION_TYPE_EXAMPLE
    For Each step In example.steps
        err_msg = vbNullString
        previous_step_status = step_status
        step_result = step_runner.run_step(step, context)
        step_status = CStr(step_result(0))
        If previous_step_status <> STATUS_OK And previous_step_status <> vbNullString Then
            ' mark all steps as skipped when the last step wasn't OK
            If step_status <> STATUS_MISSING Then step_status = STATUS_SKIPPED
        End If
        If Not step_status = STATUS_OK And Not step_status = STATUS_SKIPPED Then
            'step failed -> get the failure message
            err_msg = step_result(1)
            'this is a hack because of the broken exception handling in vba (e.g. vba will overwrite the custom err descriptions)
            If step_status = STATUS_PENDING Then err_msg = TSpec.LastFailMsg
        End If
        log_event EVENT_RUN_STEP_FINISHED, err_msg, step.OriginalStepDefinition, SECTION_TYPE_STEP, step_status
    Next
    log_event EVENT_RUN_EXAMPLE_FINISHED, example.description, example.OriginalHeadline, SECTION_TYPE_EXAMPLE
End Sub

Public Sub log_event(event_name As String, Optional event_msg, Optional section_name, Optional section_type, _
                        Optional section_status, Optional feature_file, Optional log_time)
    m_parent_runner.log_event event_name, event_msg, section_name, section_type, section_status, feature_file, log_time
End Sub

