VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' this class will listen to all events produced by Senfgurke while parsing and
' executing features and save them to EventStream collection for later inspection

Option Explicit

Private WithEvents senfgurke_event_logger As Senfgurke.TEventLogger
Attribute senfgurke_event_logger.VB_VarHelpID = -1
Private m_event_stream As Collection 'container for all consumed events

Private Sub Class_Initialize()
    Set m_event_stream = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_event_stream = Nothing
End Sub

Public Property Get SenfgurkeEventLogger() As Senfgurke.TEventLogger
    Set SenfgurkeEventLogger = senfgurke_event_logger
End Property

Public Property Set SenfgurkeEventLogger(ByVal active_logger As Senfgurke.TEventLogger)
    Set senfgurke_event_logger = active_logger
End Property

Public Property Get EventStream() As Collection
    Set EventStream = m_event_stream
End Property

Public Function filter_events(event_type As String) As Collection

    Dim log_event As Variant
    Dim matched_events As Collection
    
    Set matched_events = New Collection
    For Each log_event In m_event_stream
        If log_event.msg_type = event_type Then matched_events.Add log_event
    Next
    Set filter_events = matched_events
End Function

Private Sub senfgurke_event_logger_EventLogged(new_event As Senfgurke.TEvent)
    Me.EventStream.Add new_event
End Sub
