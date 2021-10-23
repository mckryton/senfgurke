Attribute VB_Name = "TRun"
'This is the starting point for all test runs. TRun.test will allow you to optional add tags and
' report formats while TRun.wip will call test with the @wip (work in progress) tag using the
' the verbose report format by default.

Option Explicit

Private m_step_implementations As Collection

Public Sub test(Optional filter_tag, Optional feature_filter, Optional report_format)

    Dim session As TSession
    
    Set session = THelper.new_TSession
    session.run_test StepImplementations, filter_tag, feature_filter, report_format, application_dir:=senfgurke_steps_workbook.Path
    Set session = Nothing
    Set m_step_implementations = Nothing
End Sub

Public Sub wip()
    'wip = work in progress
    test "@wip", report_format:="verbose"
End Sub

Public Sub progress(Optional filter_tag)
    test filter_tag, report_format:="progress"
End Sub

Private Property Get StepImplementations() As Collection
    
    Dim step_implementation_class As Variant

    Set m_step_implementations = New Collection
    'add all classes with step implementations here:
    For Each step_implementation_class In Array(New Steps_run_tests, New Steps_Run_Examples, New Steps_Run_Steps, _
                    New Steps_Run_features, New Steps_Load_Feature_Files, New Steps_report_verbose, _
                    New Steps_Parse_Features, New Steps_Parse_Examples, New Steps_validate_expectations, _
                    New Steps_make_steps_executable, New Steps_parse_steps, New Steps_parse_step_expressions, _
                    New Steps_parse_docstrings, New Steps_parse_tags, New Steps_parse_rules, New Steps_parse_tables, _
                    New Steps_report, New Steps_report_progress, New Steps_report_statistics, _
                    New Steps_collect_statistics, New Steps_support_functions, New Steps_save_vars_in_context, _
                    New Steps_show_step_template)
        m_step_implementations.Add step_implementation_class
    Next
    Set StepImplementations = m_step_implementations
End Property
