VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This module offers functions to create and analyze statistic events resulting
' from running features, examples and steps.

Option Explicit

Private m_step_statistics As Collection
Private m_example_statistics As Collection 'aka scenario_statistics
Private m_event_log As Collection

Public Sub count_event(event_type As String, Optional event_status As String, Optional event_name As String, Optional log_time As Long)
    
    Dim status_count As Long
    Dim event_statistics As Collection
    
    If IsMissing(event_status) Then event_status = "undefined"
    If IsMissing(event_name) Then event_name = vbNullString
    If IsMissing(log_time) Then log_time = ExtraVBA.get_unix_timestamp_now
    status_count = 1
    Select Case event_type
        Case LOG_TYPE_STEP
            Set event_statistics = step_statistics
        Case LOG_TYPE_EXAMPLE
            Set event_statistics = example_statistics
    End Select
    If ExtraVBA.collection_has_key(event_status, event_statistics) Then
        status_count = event_statistics(event_status) + 1
        event_statistics.Remove event_status
    End If
    event_statistics.Add status_count, event_status
    log_event event_type, event_name, log_time, event_status
End Sub

Public Sub log_event(log_type As String, log_msg As String, log_time As Long, Optional log_status As String)

    Dim log_event As Collection
    
    'event log will collect single log events for debugging and profiling - this is disabled by default
    If Not m_event_log Is Nothing Then
        Set log_event = New Collection
        With log_event
            .Add log_type, "type"
            .Add log_msg, "msg"
            .Add log_time, "time"
            If IsMissing(log_status) Then
                .Add vbNullString, "status"
            Else
                .Add log_status, "status"
            End If
        End With
        m_event_log.Add log_event
    End If
End Sub

Public Function filter_events(event_type As String) As Collection

    Dim log_event As Variant
    Dim matched_events As Collection
    
    Set matched_events = New Collection
    For Each log_event In m_event_log
        If log_event("type") = event_type Then matched_events.Add log_event
    Next
    Set filter_events = matched_events
End Function

Public Property Get step_statistics() As Collection
    Set step_statistics = m_step_statistics
End Property

Public Property Get example_statistics() As Collection
    Set example_statistics = m_example_statistics
End Property

Public Property Get event_log() As Collection
    Set event_log = m_event_log
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

Public Sub enable_event_log()
    'event log will collect single log events for debugging and profiling - this is disabled by default
    Set m_event_log = New Collection
End Sub

Private Sub Class_Initialize()
    Set m_step_statistics = New Collection
    Set m_example_statistics = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_step_statistics = Nothing
    Set m_example_statistics = Nothing
    Set m_event_log = Nothing
End Sub
