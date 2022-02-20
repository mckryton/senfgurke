VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TEventLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event EventLogged(new_event As TEvent)

Public Sub log_event(new_event As TEvent)
    RaiseEvent EventLogged(new_event)
End Sub

