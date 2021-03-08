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

Public Sub Given_a_feature_STR_was_reported_B5EB56508778(step_expressions As Collection)
    'Given a feature {string} was reported
    TStepVars.Session.Reporter.report REPORT_MSG_TYPE_FEATURE_NAME, "Feature: " & step_expressions(1), silent:=True
End Sub

Public Sub Given_a_rule_STR_was_reported_A5D68D23D5AE(step_expressions As Collection)
    'Given a rule {string} was reported
    TStepVars.Session.Reporter.report REPORT_MSG_TYPE_RULE, "Rule: " & step_expressions(1), silent:=True
End Sub

Public Sub Given_an_example_STR_was_reported_9C2B7078D5AE(step_expressions As Collection)
    'Given an example {string} was reported
    TStepVars.Session.Reporter.report REPORT_MSG_TYPE_EXAMPLE_TITLE, "Example: " & step_expressions(1), silent:=True
End Sub

Public Sub Given_a_description_line_STR_6B57E8EB7A49(step_expressions As Collection)
    'Given a description line {string}
    TStepVars.gherkin_text = step_expressions(1)
End Sub

Public Sub Given_a_report_message_STR_with_the_status_STR_fol_D8320E1F6BB3(step_expressions As Collection)
    'Given a report message {string} with the status {string} followed by a docstring {string}
    Dim docstring_step As String
    
    docstring_step = step_expressions(1) & vbLf _
                        & step_expressions(3)
    Set TStepVars.ReportMsg = TRun.Session.Reporter.build_msg_package(REPORT_MSG_TYPE_STEP, docstring_step, step_expressions(2))
End Sub

Public Sub When_the_execution_of_the_first_example_step_is_re_7D6F60A0AA78()
    'When the execution of the first example step is reported
    TStepVars.Session.Reporter.report REPORT_MSG_TYPE_STEP, "Given sample step", STATUS_OK, silent:=True
End Sub

Public Sub When_the_description_line_is_formatted_417E5E95BA43()
    'When the description line is formatted
    TStepVars.ReportOutput = TStepVars.Session.Reporter.Report_Formatter.format_feature_description(TStepVars.gherkin_text)
End Sub

Public Sub Then_the_report_output_contains_STR_5A8ED3F2F3C6(step_expressions As Collection)
    'Then the report output contains {string}
    TSpec.expect(TStepVars.Session.Reporter.ConsoleOutput).include step_expressions(1)
End Sub

Public Sub Then_the_report_output_doesnt_contain_STR_77804E212494(step_expressions As Collection)
    'Then the report output doesn't contain {string}
    TSpec.expect(TStepVars.Session.Reporter.ConsoleOutput).not_to_include step_expressions(1)
End Sub

Public Sub Then_the_description_is_reported_as_STR_CA847D5BC6A5(step_expressions As Collection)
    'Then the description is reported as {string}
    TSpec.expect(TStepVars.ReportOutput).to_be step_expressions(1)
End Sub

Public Sub Then_line_INT_of_the_resulting_output_is_STR_A434511445AD(step_expressions As Collection)
    'Then line {integer} of the resulting output is {string}
    Dim lines As Variant
    
    lines = Split(TStepVars.ReportOutput, vbLf)
    TSpec.expect(lines(step_expressions(1) - 1)).to_be step_expressions(2)
End Sub
