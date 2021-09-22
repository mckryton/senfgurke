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

Public Sub Given_the_report_format_is_STR_94DE9139C9BF(example_context As TContext, step_expressions As Collection)
    'Given the report format is {string}
    example_context.set_value New TSession, "session"
    example_context.get_value("session").set_report_format step_expressions(1)
    example_context.get_value("session").Reporter.enable_silent_mode
End Sub

Public Sub Given_a_code_template_for_a_missing_step_was_repor_E4A1AA906A81(example_context As TContext, step_expressions As Collection)
    'Given a code template for a missing step was reported as {string}
    example_context.get_value("session").Reporter.report REPORT_MSG_TYPE_CODE_TEMPLATE, step_expressions(1)
End Sub

Public Sub Given_a_step_STR_was_reported_with_status_STR_and__6DDD4DF1A06C(example_context As TContext, step_expressions As Collection)
    'Given a step {string} was reported with status {string} and error msg {string}
    example_context.get_value("session").Reporter.report REPORT_MSG_TYPE_STEP, step_expressions(1), step_expressions(2), step_expressions(3)
End Sub

Public Sub Given_a_step_STR_was_reported_with_status_STR_4E1E2C095DFA(example_context As TContext, step_expressions As Collection)
    'And a step {string} was reported with status {string}
    example_context.get_value("session").Reporter.report REPORT_MSG_TYPE_STEP, step_expressions(1), step_expressions(2)
End Sub

Public Sub Given_a_parse_error_STR_FF6B36833425(example_context As TContext, step_expressions As Collection)
    'Given a parse error {string}
    example_context.set_value New Collection, "parse_error"
    example_context.get_value("parse_error").Add step_expressions(1), "err_msg"
End Sub

Public Sub Given_the_error_description_is_STR_73D44E8E84D9(example_context As TContext, step_expressions As Collection)
    'And the error description is {string}
    example_context.get_value("parse_error").Add step_expressions(1), "err_description"
End Sub

Public Sub When_the_parse_error_is_reported_40CCF87A8732(example_context As TContext)
    'When the parse error is reported
    Dim parse_error As Collection
    Dim session As TSession
    
    Set parse_error = example_context.get_value("parse_error")
    Set session = example_context.get_value("session")
    session.Reporter.report REPORT_MSG_TYPE_PARSE_ERR, parse_error("err_msg") _
                                    , err_msg:=parse_error("err_description")
End Sub

Public Sub When_a_step_STR_is_reported_with_status_STR_22E595BD5DFA(example_context As TContext, step_expressions As Collection)
    'When a step {string} is reported with status {string}
    example_context.get_value("session").Reporter.report REPORT_MSG_TYPE_STEP, step_expressions(1), step_expressions(2)
End Sub

Public Sub When_a_step_STR_is_reported_with_status_STR_and_er_1F229896A06C(example_context As TContext, step_expressions As Collection)
    'When a step {string} is reported with status {string} and error msg {string}
    example_context.get_value("session").Reporter.report REPORT_MSG_TYPE_STEP, step_expressions(1), step_expressions(2), step_expressions(3)
End Sub

Public Sub When_the_report_will_report_code_templates_for_the_B3C6D3155FBC(example_context As TContext)
    'When the report will report code templates for the missing steps
    example_context.get_value("session").Reporter.finish_report
End Sub

Public Sub When_all_steps_were_reported_and_the_report_is_fin_A68058ECE13A(example_context As TContext)
    'When all steps were reported and the report is finished
    example_context.get_value("session").Reporter.finish_report
    example_context.set_value example_context.get_value("session").Reporter.ConsoleOutput, "report_output"
End Sub

Public Sub Then_the_resulting_report_output_is_STR_4C01A19A67C4(example_context As TContext, step_expressions As Collection)
    'Then the resulting report output is {string}
    Dim output As String
    
    output = ExtraVBA.trim_linebreaks(example_context.get_value("session").Reporter.ConsoleOutput)
    'TSpec.expect(example_context.get_value("session").Reporter.ConsoleOutput).to_be step_expressions(1)
    TSpec.expect(output).to_be step_expressions(1)
End Sub

Public Sub Then_the_trimmed_report_output_is_STR_BBF5E2C067C4(example_context As TContext, step_expressions As Collection)
    'Then the trimmed report output is {string}
    TSpec.expect(ExtraVBA.trim_linebreaks(example_context.get_value("session").Reporter.ConsoleOutput)).to_be step_expressions(1)
End Sub

