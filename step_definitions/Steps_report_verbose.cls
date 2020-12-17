VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_report_verbose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_report_output As String
Dim m_report_msg As Collection
Dim m_verbose_formatter As TReport_Formatter_Verbose

Public Sub Given_the_report_format_is_set_to_verbose_558493C4C322()
    'Given the report format is set to verbose
    Set m_verbose_formatter = New TReport_Formatter_Verbose
End Sub

Public Sub Given_a_feature_STR_was_reported_B5EB56508778(step_expressions As Collection)
    'Given a feature {string} was reported

    Dim report_msg As Collection
    
    Set report_msg = TReport.build_msg_package(TYPE_FEATURE_NAME, "Feature: " & step_expressions(1), vbNullString, vbNullString)
    m_verbose_formatter.process_msg report_msg, silent:=True
End Sub

Public Sub Given_a_rule_STR_was_reported_A5D68D23D5AE(step_expressions As Collection)
    'Given a rule {string} was reported

    Dim report_msg As Collection
    
    Set report_msg = TReport.build_msg_package(TYPE_RULE, "Rule: " & step_expressions(1), vbNullString, vbNullString)
    m_verbose_formatter.process_msg report_msg, silent:=True
End Sub

Public Sub Given_an_example_STR_was_reported_9C2B7078D5AE(step_expressions As Collection)
    'Given an example {string} was reported

    Dim report_msg As Collection
    
    Set report_msg = TReport.build_msg_package(TYPE_EXAMPLE_TITLE, "Example: " & step_expressions(1), vbNullString, vbNullString)
    m_verbose_formatter.process_msg report_msg, silent:=True
End Sub

Public Sub Given_a_line_from_the_feature_is_reported_as_featu_AB794F4783C4(step_expressions As Collection)
    'Given a line from the feature is reported as feature description {string}
    Set m_report_msg = TReport.build_msg_package(TYPE_DESC, step_expressions(1), vbNullString, vbNullString)
End Sub

Public Sub Given_a_feature_with_a_description_STR_F0A63D8946A5(step_expressions As Collection)
    'Given a feature with a description {string}
    Set m_report_msg = TReport.build_msg_package(TYPE_DESC, step_expressions(1), vbNullString, vbNullString)
End Sub

Public Sub Given_a_step_STR_with_the_status_STR_FAF1B1E8C8DC(step_expressions As Collection)
    'Given a step {string} with the status {string}
    Set m_report_msg = TReport.build_msg_package(TYPE_STEP, step_expressions(1), step_expressions(2), vbNullString)
End Sub

Public Sub Given_a_step_STR_followed_by_a_docstring_STR_with__B7F15CFA718B(step_expressions As Collection)
    'Given a step {string} followed by a docstring {string} with the status {string}
    Dim docstring_step As String
    
    docstring_step = step_expressions(1) & vbLf _
                        & """""""" & vbLf _
                        & step_expressions(2) & vbLf _
                        & """"""""
    Set m_report_msg = TReport.build_msg_package(TYPE_STEP, docstring_step, step_expressions(3), vbNullString)
End Sub

 Public Sub Given_a_rule_STR_F5C6A47A4718(step_expressions As Collection)
    'Given a rule {string}
    Set m_report_msg = TReport.build_msg_package(TYPE_RULE, Replace(step_expressions(1), "<br>", vbLf))
End Sub

Public Sub Given_a_step_STR_fails_with_the_error_message_STR_F54FA807B797(step_expressions As Collection)
    'Given a step {string} fails with the error message {string}
    Set m_report_msg = TReport.build_msg_package(TYPE_STEP, step_expressions(1), STATUS_FAIL, step_expressions(2))
End Sub

Public Sub When_the_execution_of_the_first_example_step_is_re_7D6F60A0AA78()
    'When the execution of the first example step is reported
    
    Dim report_msg As Collection
    
    Set report_msg = TReport.build_msg_package(TYPE_STEP, "Given sample step", STATUS_OK, vbNullString)
    m_report_output = m_verbose_formatter.process_msg(report_msg, silent:=True)
End Sub

Public Sub When_the_reported_message_is_prepared_as_output_fo_5E1841FB3AC8()
    'When the reported message is prepared as output for the report
    m_report_output = m_verbose_formatter.prepare_msg(m_report_msg)
End Sub

Public Sub Then_the_report_output_contains_STR_5A8ED3F2F3C6(step_expressions As Collection)
    'Then the report output contains {string}
    TSpec.expect(m_report_output).include step_expressions(1)
End Sub

Public Sub Then_the_report_output_doesnt_contain_STR_77804E212494(step_expressions As Collection)
    'Then the report output doesn't contain {string}
    TSpec.expect(m_report_output).not_to_include step_expressions(1)
End Sub

Public Sub Then_the_description_is_reported_as_STR_CA847D5BC6A5(step_expressions As Collection)
    'Then the description is reported as {string}
    TSpec.expect(m_report_output).to_be step_expressions(1)
End Sub

Public Sub Then_the_resulting_output_is_set_to_STR_4C01E512CDA4(step_expressions As Collection)
    'Then the resulting output is set to {string}
    TSpec.expect(m_report_output).to_be step_expressions(1)
End Sub

Public Sub Then_the_first_line_of_the_resulting_output_is_set_2D1D1353FBA6(step_expressions As Collection)
    'Then the first line of the resulting output is set to {string}
    TSpec.expect(Split(m_report_output, vbLf)(0)).to_be step_expressions(1)
End Sub

Public Sub Then_2nd_and_4th_line_are_STR_A56F9A13380D(step_expressions As Collection)
    'And 2nd and 4th line are {string}
    TSpec.expect(Split(m_report_output, vbLf)(1)).to_be step_expressions(1)
    TSpec.expect(Split(m_report_output, vbLf)(3)).to_be step_expressions(1)
End Sub

Public Sub Then_the_3rd_line_is_STR_DBF0DB00D378(step_expressions As Collection)
    'And the 3rd line is {string}
    TSpec.expect(Split(m_report_output, vbLf)(2)).to_be step_expressions(1)
End Sub

Public Sub Then_the_resulting_output_starts_with_a_new_line_3342859E2E9A()
    'Then the resulting output starts with a new line
    TSpec.expect(Left(m_report_output, 1)).to_be vbLf
End Sub

Public Sub Then_every_line_after_the_first_one_is_indented_by_D63F5EC8025D(step_expressions As Collection)
    'Then every line after the first one is indented by 7 spaces
    TSpec.expect(Left(Split(m_report_output, vbLf)(2), 7)).to_be Space$(step_expressions(1))
End Sub

Public Sub Then_the_second_line_shows_the_indented_error_mess_5F3E66D58D32()
    'Then the second line shows the indented error message
    TSpec.expect(Split(m_report_output, vbLf)(1)).to_be Space$(10) & "err: sample err msg"
End Sub
