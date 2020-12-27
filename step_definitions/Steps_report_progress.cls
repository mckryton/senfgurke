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

Public Sub Given_the_report_format_is_set_to_progress_4F152AAE8E9C()
    'And the report format is set to progress
    Set TStepVars.ReportFormatter = New TReportFormatterProgress
End Sub

Public Sub Given__INT_steps_were_reported_as_successful_53AEBD4A25F4(step_expressions As Collection)
    'Given  {integer} steps were reported as successful
    
    Dim index As Integer
    
    Set TStepVars.ReportMessages = New Collection
    For index = 1 To step_expressions(1)
        TStepVars.ReportMessages.Add TRun.Reporter.build_msg_package(REPORT_MSG_TYPE_STEP, "Given a successful step", "OK", vbNullString)
    Next
End Sub

Public Sub Given_the_next_step_STR_was_reported_as_missing_07941E95CA02(step_expressions As Collection)
    'And the next step {string} was reported as missing
    TStepVars.ReportMessages.Add TRun.Reporter.build_msg_package(REPORT_MSG_TYPE_STEP, step_expressions(1), "MISSING", vbNullString)
End Sub
