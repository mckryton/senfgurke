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


Public Sub Given_a_report_message_about_a_STR_with_the_status_F720FF76FEB9(step_expressions As Collection)
    'Given a report message about a {string} with the status {string}
    Set TStepVars.ReportMsg = TRun.Session.Reporter.build_msg_package(REPORT_MSG_TYPE_STEP, step_expressions(1), step_expressions(2), vbNullString)
End Sub

Public Sub Given_a_report_message_STR_with_status_STR_and_err_9EBC51E6C63C(step_expressions As Collection)
    'Given a report message {string} with status {string} and error msg {string}
    Set TStepVars.ReportMsg = TRun.Session.Reporter.build_msg_package(REPORT_MSG_TYPE_STEP, step_expressions(1), step_expressions(2), step_expressions(3))
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

Public Sub When_the_reported_message_is_formatted_84DF0DEABA43()
    'When the reported message is formatted
    TStepVars.ReportOutput = TStepVars.Session.Reporter.Report_Formatter.format_msg_content(TStepVars.ReportMsg)
End Sub

Public Sub When_step_results_are_reported_FC23061D41C3()
    'When step results are reported
    
    Dim report_msg As Collection
    Dim output As String
    
    TStepVars.ReportOutput = vbNullString
    For Each report_msg In TStepVars.ReportMessages
        output = TStepVars.ReportFormatter.process_msg(report_msg, silent:=True)
        TStepVars.ReportOutput = TStepVars.ReportOutput & output
    Next
End Sub

Public Sub When_the_report_will_report_code_templates_for_the_B3C6D3155FBC()
    'When the report will report code templates for the missing steps
    TStepVars.ReportOutput = TStepVars.Session.Reporter.Report_Formatter.get_code_templates_for_missing_steps
End Sub

Public Sub When_all_steps_were_reported_and_the_report_is_fin_A68058ECE13A()
    'When all steps were reported and the report is finished
    TStepVars.Session.Reporter.finish_report silent:=True
    TStepVars.ReportOutput = TStepVars.Session.Reporter.ConsoleOutput
End Sub

Public Sub Then_the_resulting_report_output_is_STR_4C01A19A67C4(step_expressions As Collection)
    'Then the resulting report output is {string}
    TSpec.expect(TStepVars.ReportOutput).to_be step_expressions(1)
End Sub


