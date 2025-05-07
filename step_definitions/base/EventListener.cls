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

Public WithEvents SenfgurkeEventLogger As Senfgurke.TEventLogger
Attribute SenfgurkeEventLogger.VB_VarHelpID = -1
Public EventStream As Collection  'container for all consumed events

Private Sub Class_Initialize()
    Set EventStream = New Collection
End Sub

Private Sub Class_Terminate()
    Set EventStream = Nothing
End Sub

Public Function filter_events(event_name As String) As Collection

    Dim log_event As Variant
    Dim matched_events As Collection
    
    Set matched_events = New Collection
    For Each log_event In Me.EventStream
        If log_event.EventName = event_name Then matched_events.Add log_event
    Next
    Set filter_events = matched_events
End Function

Private Sub SenfgurkeEventLogger_EventLogged(new_event As Senfgurke.TEvent)
    Me.EventStream.Add new_event
End Sub
