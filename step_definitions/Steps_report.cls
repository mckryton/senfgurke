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
    Dim session As TSession
    
    Set session = THelper.new_TSession
    With session
        .set_report_format step_expressions(1)
        .reporter.enable_silent_mode
    End With
    example_context.set_value session, "session"
End Sub

Public Sub Given_a_code_template_for_a_missing_step_was_repor_E4A1AA906A81(example_context As TContext, step_expressions As Collection)
    'Given a code template for a missing step was reported as {string}
    Dim session As TSession
    
    Set session = example_context.get_value("session")
    session.log_event EVENT_STEP_IS_MISSING, step_expressions(1)
End Sub

Public Sub Given_steps_were_reported_as_5E283B73EF13(example_context As TContext, data_table As TDataTable)
    'Given steps were reported as
    Dim row As Variant
    Dim session As TSession
    
    Set session = example_context.get_value("session")
    For Each row In data_table.table_rows
        session.log_event EVENT_RUN_STEP_FINISHED, row("err_msg"), CStr(row("step_name")), SECTION_TYPE_STEP, row("status")
    Next
End Sub

Public Sub Given_a_parse_error_STR_was_found_in_STR_570BB7623FA8(example_context As TContext, step_expressions As Collection)
    'Given a parse error {string} was found in {string}
    Dim parse_error As Collection
    
    Set parse_error = New Collection
    parse_error.Add step_expressions(1), "err_description"
    parse_error.Add step_expressions(2), "feature_file"
    example_context.set_value parse_error, "parse_error"
End Sub

Public Sub When_the_parse_error_is_reported_40CCF87A8732(example_context As TContext)
    'When the parse error is reported
    Dim parse_error As Collection
    Dim session As TSession
    
    Set parse_error = example_context.get_value("parse_error")
    Set session = example_context.get_value("session")
    session.log_event EVENT_PARSE_ERROR, parse_error("err_description"), parse_error("feature_file"), SECTION_TYPE_FEATURE
End Sub

Public Sub When_a_step_STR_is_reported_with_status_STR_22E595BD5DFA(example_context As TContext, step_expressions As Collection)
    'When a step {string} is reported with status {string}
    Dim session As TSession
    
    Set session = example_context.get_value("session")
    session.log_event EVENT_RUN_STEP_FINISHED, vbNullString, step_expressions(1), SECTION_TYPE_STEP, step_expressions(2)
End Sub

Public Sub When_a_step_STR_is_reported_with_status_STR_and_er_1F229896A06C(example_context As TContext, step_expressions As Collection)
    'When a step {string} is reported with status {string} and error msg {string}
    Dim session As TSession
    
    Set session = example_context.get_value("session")
    session.log_event EVENT_RUN_STEP_FINISHED, step_expressions(3), step_expressions(1), SECTION_TYPE_STEP, step_expressions(2)
End Sub

Public Sub When_the_report_will_report_code_templates_for_the_B3C6D3155FBC(example_context As TContext)
    'When the report will report code templates for the missing steps
    Dim session As TSession
    
    Set session = example_context.get_value("session")
    session.reporter.finish_report
End Sub

Public Sub When_all_steps_were_reported_and_the_report_is_fin_A68058ECE13A(example_context As TContext)
    'When all steps were reported and the report is finished
    example_context.get_value("session").reporter.finish_report
    example_context.set_value example_context.get_value("session").reporter.ConsoleOutput, "report_output"
End Sub

Public Sub Then_the_resulting_report_output_is_STR_4C01A19A67C4(example_context As TContext, step_expressions As Collection)
    'Then the resulting report output is {string}
    Dim output As String
    
    output = ExtraVBA.trim_linebreaks(example_context.get_value("session").reporter.ConsoleOutput)
    TSpec.expect(output).to_be step_expressions(1)
End Sub

Public Sub Then_the_trimmed_report_output_is_STR_BBF5E2C067C4(example_context As TContext, step_expressions As Collection)
    'Then the trimmed report output is {string}
    TSpec.expect(ExtraVBA.trim_linebreaks(example_context.get_value("session").reporter.ConsoleOutput)).to_be step_expressions(1)
End Sub

