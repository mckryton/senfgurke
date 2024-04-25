VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TExecutionHooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'After event hooks are triggered by selected log events and the causing log event is
' used as a parameter for the after event hook function, eg. to read the status of
' the causing event

Event AfterExampleHook(new_event As TEvent)

Public Sub call_after_example_hook(new_event As TEvent)
    RaiseEvent AfterExampleHook(new_event)
End Sub

