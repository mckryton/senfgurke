VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_the_report_format_is_STR_94DE9139C9BF(step_expressions As Collection)
    'Given the report format is {string}
    Set TStepVars.Session = New TSession
    TStepVars.Session.set_report_format step_expressions(1)
    Set TStepVars.ReportFormatter = TStepVars.Session.Reporter.Report_Formatter
End Sub

Public Sub Given_a_code_template_for_a_missing_step_was_repor_E4A1AA906A81(step_expressions As Collection)
    'Given a code template for a missing step was reported as {string}
    TStepVars.CodeTemplate = step_expressions(1)
    TStepVars.Session.Reporter.report REPORT_MSG_TYPE_CODE_TEMPLATE, TStepVars.CodeTemplate, silent:=True
End Sub

Public Sub Given_a_step_STR_was_reported_with_status_STR_and__6DDD4DF1A06C(step_expressions As Collection)
    'Given a step {string} was reported with status {string} and error msg {string}
    TStepVars.Session.Reporter.report REPORT_MSG_TYPE_STEP, step_expressions(1), step_expressions(2), step_expressions(3), silent:=True
End Sub

Public Sub Given_a_step_STR_was_reported_with_status_STR_4E1E2C095DFA(step_expressions As Collection)
    'And a step {string} was reported with status {string}
    TStepVars.Session.Reporter.report REPORT_MSG_TYPE_STEP, step_expressions(1), step_expressions(2), silent:=True
End Sub

Public Sub Given_a_parse_error_STR_FF6B36833425(step_expressions As Collection)
    'Given a parse error {string}
    Set TStepVars.parse_error = New Collection
    TStepVars.parse_error.Add step_expressions(1), "err_msg"
End Sub

Public Sub Given_the_error_description_is_STR_73D44E8E84D9(step_expressions As Collection)
    'And the error description is {string}
    TStepVars.parse_error.Add step_expressions(1), "err_description"
End Sub

Public Sub When_the_parse_error_is_reported_40CCF87A8732()
    'When the parse error is reported
    TStepVars.Session.Reporter.report REPORT_MSG_TYPE_PARSE_ERR, TStepVars.parse_error("err_msg"), err_msg:=TStepVars.parse_error("err_description"), silent:=True
End Sub

Public Sub When_a_step_STR_is_reported_with_status_STR_22E595BD5DFA(step_expressions As Collection)
    'When a step {string} is reported with status {string}
    TStepVars.Session.Reporter.report REPORT_MSG_TYPE_STEP, step_expressions(1), step_expressions(2), silent:=True
End Sub

Public Sub When_a_step_STR_is_reported_with_status_STR_and_er_1F229896A06C(step_expressions As Collection)
    'When a step {string} is reported with status {string} and error msg {string}
    TStepVars.Session.Reporter.report REPORT_MSG_TYPE_STEP, step_expressions(1), step_expressions(2), step_expressions(3), silent:=True
End Sub

Public Sub When__INT_steps_are_reported_as_successful_34289162846E(step_expressions As Collection)
    'When  {integer} steps are reported as successful
    Dim index As Integer
    
    Set TStepVars.ReportMessages = New Collection
    For index = 1 To step_expressions(1)
        TStepVars.Session.Reporter.report REPORT_MSG_TYPE_STEP, "Given a valid step", STATUS_OK, silent:=True
    Next
End Sub

Public Sub When_the_report_will_report_code_templates_for_the_B3C6D3155FBC()
    'When the report will report code templates for the missing steps
    TStepVars.Session.Reporter.finish_report silent:=True
End Sub

Public Sub When_all_steps_were_reported_and_the_report_is_fin_A68058ECE13A()
    'When all steps were reported and the report is finished
    TStepVars.Session.Reporter.finish_report silent:=True
    TStepVars.ReportOutput = TStepVars.Session.Reporter.ConsoleOutput
End Sub

Public Sub Then_the_resulting_report_output_is_STR_4C01A19A67C4(step_expressions As Collection)
    'Then the resulting report output is {string}
    TSpec.expect(TStepVars.Session.Reporter.ConsoleOutput).to_be step_expressions(1)
End Sub

Public Sub Then_the_trimmed_report_output_is_STR_BBF5E2C067C4(step_expressions As Collection)
    'Then the trimmed report output is {string}
    TSpec.expect(ExtraVBA.trim_linebreaks(TStepVars.Session.Reporter.ConsoleOutput)).to_be step_expressions(1)
End Sub

