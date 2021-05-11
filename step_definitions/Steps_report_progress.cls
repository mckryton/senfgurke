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

Public Sub Given_the_next_step_STR_was_reported_as_missing_07941E95CA02(context As TContext, step_expressions As Collection)
    'And the next step {string} was reported as missing
    TStepVars.ReportMessages.Add TRun.session.Reporter.build_msg_package(REPORT_MSG_TYPE_STEP, step_expressions(1), "MISSING", vbNullString)
End Sub
