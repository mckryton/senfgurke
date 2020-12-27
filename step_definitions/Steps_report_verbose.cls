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

Public Sub Given_the_report_format_is_set_to_verbose_558493C4C322()
    'Given the report format is set to verbose
    Set TStepVars.ReportFormatter = New TReportFormatterVerbose
End Sub

Public Sub Given_a_feature_STR_was_reported_B5EB56508778(step_expressions As Collection)
    'Given a feature {string} was reported

    Dim report_msg As Collection
    
    Set report_msg = TRun.Reporter.build_msg_package(REPORT_MSG_TYPE_FEATURE_NAME, "Feature: " & step_expressions(1), vbNullString, vbNullString)
    TStepVars.ReportFormatter.process_msg report_msg, silent:=True
End Sub

Public Sub Given_a_rule_STR_was_reported_A5D68D23D5AE(step_expressions As Collection)
    'Given a rule {string} was reported

    Dim report_msg As Collection
    
    Set report_msg = TRun.Reporter.build_msg_package(REPORT_MSG_TYPE_RULE, "Rule: " & step_expressions(1), vbNullString, vbNullString)
    TStepVars.ReportFormatter.process_msg report_msg, silent:=True
End Sub

Public Sub Given_an_example_STR_was_reported_9C2B7078D5AE(step_expressions As Collection)
    'Given an example {string} was reported

    Dim report_msg As Collection
    
    Set report_msg = TRun.Reporter.build_msg_package(REPORT_MSG_TYPE_EXAMPLE_TITLE, "Example: " & step_expressions(1), vbNullString, vbNullString)
    TStepVars.ReportFormatter.process_msg report_msg, silent:=True
End Sub

Public Sub Given_a_line_from_the_feature_is_reported_as_featu_AB794F4783C4(step_expressions As Collection)
    'Given a line from the feature is reported as feature description {string}
    Set TStepVars.ReportMsg = TRun.Reporter.build_msg_package(REPORT_MSG_TYPE_DESC, step_expressions(1), vbNullString, vbNullString)
End Sub

Public Sub Given_a_feature_with_a_description_STR_F0A63D8946A5(step_expressions As Collection)
    'Given a feature with a description {string}
    Set TStepVars.ReportMsg = TRun.Reporter.build_msg_package(REPORT_MSG_TYPE_DESC, step_expressions(1), vbNullString, vbNullString)
End Sub

Public Sub Given_a_step_STR_followed_by_a_docstring_STR_with__B7F15CFA718B(step_expressions As Collection)
    'Given a step {string} followed by a docstring {string} with the status {string}
    Dim docstring_step As String
    
    docstring_step = step_expressions(1) & vbLf _
                        & """""""" & vbLf _
                        & step_expressions(2) & vbLf _
                        & """"""""
    Set TStepVars.ReportMsg = TRun.Reporter.build_msg_package(REPORT_MSG_TYPE_STEP, docstring_step, step_expressions(3), vbNullString)
End Sub

 Public Sub Given_a_rule_STR_F5C6A47A4718(step_expressions As Collection)
    'Given a rule {string}
    Set TStepVars.ReportMsg = TRun.Reporter.build_msg_package(REPORT_MSG_TYPE_RULE, Replace(step_expressions(1), "<br>", vbLf))
End Sub

Public Sub Given_a_step_STR_fails_with_the_error_message_STR_F54FA807B797(step_expressions As Collection)
    'Given a step {string} fails with the error message {string}
    Set TStepVars.ReportMsg = TRun.Reporter.build_msg_package(REPORT_MSG_TYPE_STEP, step_expressions(1), STATUS_FAIL, step_expressions(2))
End Sub

Public Sub When_the_execution_of_the_first_example_step_is_re_7D6F60A0AA78()
    'When the execution of the first example step is reported
    
    Dim report_msg As Collection
    
    Set report_msg = TRun.Reporter.build_msg_package(REPORT_MSG_TYPE_STEP, "Given sample step", STATUS_OK, vbNullString)
    TStepVars.ReportOutput = TStepVars.ReportFormatter.process_msg(report_msg, silent:=True)
End Sub

Public Sub Then_the_report_output_contains_STR_5A8ED3F2F3C6(step_expressions As Collection)
    'Then the report output contains {string}
    TSpec.expect(TStepVars.ReportOutput).include step_expressions(1)
End Sub

Public Sub Then_the_report_output_doesnt_contain_STR_77804E212494(step_expressions As Collection)
    'Then the report output doesn't contain {string}
    TSpec.expect(TStepVars.ReportOutput).not_to_include step_expressions(1)
End Sub

Public Sub Then_the_description_is_reported_as_STR_CA847D5BC6A5(step_expressions As Collection)
    'Then the description is reported as {string}
    TSpec.expect(TStepVars.ReportOutput).to_be step_expressions(1)
End Sub

Public Sub Then_the_resulting_output_is_set_to_STR_4C01E512CDA4(step_expressions As Collection)
    'Then the resulting output is set to {string}
    TSpec.expect(TStepVars.ReportOutput).to_be step_expressions(1)
End Sub

Public Sub Then_the_first_line_of_the_resulting_output_is_set_2D1D1353FBA6(step_expressions As Collection)
    'Then the first line of the resulting output is set to {string}
    TSpec.expect(Split(TStepVars.ReportOutput, vbLf)(0)).to_be step_expressions(1)
End Sub

Public Sub Then_2nd_and_4th_line_are_STR_A56F9A13380D(step_expressions As Collection)
    'And 2nd and 4th line are {string}
    TSpec.expect(Split(TStepVars.ReportOutput, vbLf)(1)).to_be step_expressions(1)
    TSpec.expect(Split(TStepVars.ReportOutput, vbLf)(3)).to_be step_expressions(1)
End Sub

Public Sub Then_the_3rd_line_is_STR_DBF0DB00D378(step_expressions As Collection)
    'And the 3rd line is {string}
    TSpec.expect(Split(TStepVars.ReportOutput, vbLf)(2)).to_be step_expressions(1)
End Sub

Public Sub Then_the_resulting_output_starts_with_a_new_line_3342859E2E9A()
    'Then the resulting output starts with a new line
    TSpec.expect(Left(TStepVars.ReportOutput, 1)).to_be vbLf
End Sub

Public Sub Then_every_line_after_the_first_one_is_indented_by_D63F5EC8025D(step_expressions As Collection)
    'Then every line after the first one is indented by 7 spaces
    TSpec.expect(Left(Split(TStepVars.ReportOutput, vbLf)(2), 7)).to_be Space$(step_expressions(1))
End Sub

Public Sub Then_the_second_line_shows_the_indented_error_mess_5F3E66D58D32()
    'Then the second line shows the indented error message
    TSpec.expect(Split(TStepVars.ReportOutput, vbLf)(1)).to_be Space$(10) & "err: sample err msg"
End Sub
