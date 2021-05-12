VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_report_progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub When__INT_steps_are_reported_as_successful_34289162846E(example_context As TContext, step_expressions As Collection)
    'When  {integer} steps are reported as successful
    Dim index As Integer
    
    For index = 1 To step_expressions(1)
        example_context.get_value("session").Reporter.report REPORT_MSG_TYPE_STEP, "Given a valid step", STATUS_OK
    Next
End Sub
