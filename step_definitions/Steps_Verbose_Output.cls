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

Public Sub Given_a_step_this_is_a_step_with_the_status_OK_0E21472CE4EB()
    'Given a step "this is a step" with the status "OK"
    Set m_report_msg = TReport.build_msg_package(TYPE_STEP, "this is a step", "OK", vbNullString)
End Sub

Public Sub Given_a_step_this_is_a_step_with_the_status_FAIL_360E08B64EC7()
    'Given a step "this is a step" with the status "FAIL"
    Set m_report_msg = TReport.build_msg_package(TYPE_STEP, "this is a step", "FAIL", vbNullString)
End Sub

Public Sub Given_a_step_this_is_a_step_with_the_status_MISSIN_21F60E25EEB6()
    'Given a step "this is a step" with the status "MISSING"
    Set m_report_msg = TReport.build_msg_package(TYPE_STEP, "this is a step", "MISSING", vbNullString)
End Sub

Public Sub Given_a_step_this_is_a_step_with_the_status_PENDIN_21F60E25F5CD()
    'Given a step "this is a step" with the status "PENDING"
    Set m_report_msg = TReport.build_msg_package(TYPE_STEP, "this is a step", "PENDING", vbNullString)
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

Public Sub Then_the_resulting_output_is_set_to______OK________797C1D1BB64A()
    'Then the resulting output is set to "     OK       this is a step"
    TSpec.expect(m_report_output).to_be "     OK       this is a step"
End Sub

Public Sub Then_the_resulting_output_is_set_to______FAIL______797C4A9D98D0()
    'Then the resulting output is set to "     FAIL     this is a step"
    TSpec.expect(m_report_output).to_be "     FAIL     this is a step"
End Sub

Public Sub Then_the_resulting_output_is_set_to______MISSING___797CBC9B3E08()
    'Then the resulting output is set to "     MISSING  this is a step"
    TSpec.expect(m_report_output).to_be "     MISSING  this is a step"
End Sub

Public Sub Then_the_resulting_output_is_set_to______PENDING___797CE9928C61()
    'Then the resulting output is set to "     PENDING  this is a step"
    TSpec.expect(m_report_output).to_be "     PENDING  this is a step"
End Sub
