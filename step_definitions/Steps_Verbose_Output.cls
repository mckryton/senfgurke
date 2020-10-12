VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Verbose_Output"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_report_output As String
Dim m_report_msg As Collection
Dim m_verbose_formatter As TReport_Formatter_Verbose

Public Sub Given_the_report_format_is_set_to_verbose_E5DA50C752C0()
    'Given the report format is set to verbose
    Set m_verbose_formatter = New TReport_Formatter_Verbose
End Sub

Public Sub Given_a_feature_sample_feature_was_reported_0D24F9E46180()
    'Given a feature "sample feature" was reported

    Dim report_msg As Collection
    
    Set report_msg = TReport.build_msg_package(TYPE_FEATURE_NAME, "Feature: sample feature", vbNullString, vbNullString)
    m_verbose_formatter.process_msg report_msg, silent:=True
End Sub

Public Sub Given_a_rule_empty_rule_was_reported_54FD3DF8F4CE()
    'Given a rule "empty rule" was reported

    Dim report_msg As Collection
    
    Set report_msg = TReport.build_msg_package(TYPE_RULE, "Rule: empty rule", vbNullString, vbNullString)
    m_verbose_formatter.process_msg report_msg, silent:=True
End Sub

Public Sub Given_a_rule_has_example_was_reported_5994B2613C44()
    'Given a rule "has example" was reported

    Dim report_msg As Collection
    
    Set report_msg = TReport.build_msg_package(TYPE_RULE, "Rule: has example", vbNullString, vbNullString)
    m_verbose_formatter.process_msg report_msg, silent:=True
End Sub

Public Sub Given_an_example_sample_example_was_reported_DADE4FE80353()
    'Given an example "sample example" was reported

    Dim report_msg As Collection
    
    Set report_msg = TReport.build_msg_package(TYPE_EXAMPLE_TITLE, "Example: sample example", vbNullString, vbNullString)
    m_verbose_formatter.process_msg report_msg, silent:=True
End Sub

Public Sub Given_a_line_from_the_feature_is_reported_as_featu_41CD67A527F0()
    'Given a line from the feature is reported as feature description "description line"
    Set m_report_msg = TReport.build_msg_package(TYPE_DESC, "description line", vbNullString, vbNullString)
End Sub

Public Sub Given_a_feature_with_a_description______________de_57503E7AB70E()
    'Given a feature with a description "             description line"
    Set m_report_msg = TReport.build_msg_package(TYPE_DESC, "description line", vbNullString, vbNullString)
End Sub

Public Sub Given_a_step_Given_a_sample_step_with_the_status_O_5B21E75AB1FF()
    'Given a step "Given a sample step" with the status "OK"
    Set m_report_msg = TReport.build_msg_package(TYPE_STEP, "Given a sample step", "OK", vbNullString)
End Sub

Public Sub Given_a_step_Given_a_sample_step_with_the_status_F_F41A25B4E7EE()
    'Given a step "Given a sample step" with the status "FAIL"
    Set m_report_msg = TReport.build_msg_package(TYPE_STEP, "Given a sample step", "FAIL", vbNullString)
End Sub

Public Sub Given_a_step_Given_a_sample_step_with_the_status_M_1374A1806E0C()
    'Given a step "Given a sample step" with the status "MISSING"
    Set m_report_msg = TReport.build_msg_package(TYPE_STEP, "Given a sample step", "MISSING", vbNullString)
End Sub

Public Sub Given_a_step_Given_a_sample_step_with_the_status_P_1374A1807577()
    'Given a step "Given a sample step" with the status "PENDING"
    Set m_report_msg = TReport.build_msg_package(TYPE_STEP, "Given a sample step", "PENDING", vbNullString)
End Sub

Public Sub Given_a_rule_this_is_a_sample_rulebrand_this_is_li_66F25EBBC7D8()
    'Given a rule "this is a sample rule<br>and this is line 2 of the same rule"
    Set m_report_msg = TReport.build_msg_package(TYPE_RULE, "this is a sample rule" & vbLf _
                                                & " and this is line 2 of the same rule", "PENDING", vbNullString)
End Sub

Public Sub Given_a_step_Given_a_sample_step_fails_with_the_er_A7846F27AC5A()
    'Given a step "Given a sample step" fails with the error message "err: sample err msg"
    Set m_report_msg = TReport.build_msg_package(TYPE_STEP, "Given a sample step", "FAIL", "err: sample err msg")
End Sub

Public Sub When_the_execution_of_the_first_example_step_is_re_D9239339DB19()
    'When the execution of the first example step is reported
    
    Dim report_msg As Collection
    
    Set report_msg = TReport.build_msg_package(TYPE_STEP, "Given sample step", "OK", vbNullString)
    m_report_output = m_verbose_formatter.process_msg(report_msg, silent:=True)
End Sub

Public Sub When_the_reported_message_is_prepared_as_output_fo_FD1CD2A06B2B()
    'When the reported message is prepared as output for the report
    m_report_output = m_verbose_formatter.prepare_msg(m_report_msg)
End Sub

Public Sub Then_the_report_output_contains_Feature_sample_fea_8DE9A6ECF9F3()
    'Then the report output contains "Feature: sample feature"
    TSpec.expect(m_report_output).include "Feature: sample feature"
End Sub

Public Sub Then_the_report_output_contains_Rule_has_example_59141C67DD91()
    'Then the report output contains "Rule: has example"
    TSpec.expect(m_report_output).include "Rule: has example"
End Sub

Public Sub Then_the_report_output_doesnt_contain_Rule_empty_r_8DE9E0B03217()
    'Then the report output doesn't contain "Rule: empty rule"
    TSpec.expect(m_report_output).not_to_include "Rule: empty rule"
End Sub

Public Sub Then_the_description_is_reported_as_____descriptio_DC415FBFDF74()
    'Then the description is reported as "    description line"
    TSpec.expect(m_report_output).to_be "    description line"
End Sub

Public Sub Then_the_resulting_output_is_set_to______OK________0579C48400A4()
    'Then the resulting output is set to "     OK       Given a sample step"
    TSpec.expect(m_report_output).to_be "     OK       Given a sample step"
End Sub

Public Sub Then_the_resulting_output_is_set_to______FAIL______0579278900A4()
    'Then the resulting output is set to "     FAIL     Given a sample step"
    TSpec.expect(m_report_output).to_be "     FAIL     Given a sample step"
End Sub

Public Sub Then_the_resulting_output_is_set_to______MISSING___05796940D9E8()
    'Then the resulting output is set to "     MISSING  Given a sample step"
    TSpec.expect(m_report_output).to_be "     MISSING  Given a sample step"
End Sub

Public Sub Then_the_resulting_output_is_set_to______PENDING___057958CFD9E8()
    'Then the resulting output is set to "     PENDING  Given a sample step"
    TSpec.expect(m_report_output).to_be "     PENDING  Given a sample step"
End Sub

Public Sub Then_the_resulting_output_starts_with_a_new_line_3C50FD6E1351()
    'Then the resulting output starts with a new line
    TSpec.expect(Left(m_report_output, 1)).to_be vbLf
End Sub

Public Sub Then_every_line_after_the_first_one_is_indented_by_B850B3BC8E58()
    'Then every line after the first one is indented by 7 spaces
    TSpec.expect(Left(Split(m_report_output, vbLf)(2), 7)).to_be Space$(7)
End Sub

Public Sub Then_the_resulting_output_starts_with_the_step_sta_D72578893FDD()
    'Then the resulting output starts with the step status and name
    TSpec.expect(Split(m_report_output, vbLf)(0)).to_be "     FAIL     Given a sample step"
End Sub

Public Sub Then_the_second_line_shows_the_indented_error_mess_629169F5560B()
    'Then the second line shows the indented error message
    TSpec.expect(Split(m_report_output, vbLf)(1)).to_be Space$(10) & "err: sample err msg"
End Sub
