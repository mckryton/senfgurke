VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This module offers functions to create and analyze statistic events resulting
' from running features, examples and steps.

Option Explicit

Private m_step_statistics As Collection
Private m_example_statistics As Collection 'aka scenario_statistics
Private WithEvents m_event_logger As TEventLogger
Attribute m_event_logger.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set m_step_statistics = New Collection
    Set m_example_statistics = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_step_statistics = Nothing
    Set m_example_statistics = Nothing
End Sub

Public Property Set active_logger(ByVal active_logger As TEventLogger)
    Set m_event_logger = active_logger
End Property

Private Sub count_event(event_name As String, event_status As String, log_time As Single)
    
    Dim status_count As Long
    Dim event_statistics As Collection
    
    status_count = 1
    Select Case event_name
        Case LOG_EVENT_RUN_STEP
            Set event_statistics = step_statistics
        Case LOG_EVENT_RUN_EXAMPLE
            Set event_statistics = example_statistics
        Case Else
            Exit Sub
    End Select
    If ExtraVBA.collection_has_key(event_status, event_statistics) Then
        status_count = event_statistics(event_status) + 1
        event_statistics.Remove event_status
    End If
    event_statistics.Add status_count, event_status
End Sub

Public Property Get step_statistics() As Collection
    Set step_statistics = m_step_statistics
End Property

Public Property Get example_statistics() As Collection
    Set example_statistics = m_example_statistics
End Property

Public Function get_statistics_summary() As String

    Dim states As String
    Dim state As Variant
    Dim state_id As String
    Dim step_sum_count As Long
    Dim step_state_count As Long
    Dim summary As String
    
    step_sum_count = 0
    summary = "("
    states = "OK:passed|FAIL:failed|MISSING:undefined|PENDING:pending|SKIPPED:skipped"
    For Each state In Split(states, "|")
        state_id = Split(state, ":")(0)
        If ExtraVBA.collection_has_key(state_id, step_statistics) Then
          step_state_count = step_statistics(state_id)
          summary = summary & step_state_count & " " & Split(state, ":")(1) & ", "
          step_sum_count = step_sum_count + step_state_count
        End If
    Next
    If step_sum_count = 0 Then
        get_statistics_summary = "0 steps"
    ElseIf step_sum_count = 1 Then
        get_statistics_summary = "1 step " & Left(summary, Len(summary) - 2) & ")"
    Else
        get_statistics_summary = step_sum_count & " steps " & Left(summary, Len(summary) - 2) & ")"
    End If
End Function

Private Sub m_event_logger_EventLogged(new_event As TEvent)
    count_event new_event.EventName, new_event.SectionStatus, new_event.LogTime
End Sub
