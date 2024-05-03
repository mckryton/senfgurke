Attribute VB_Name = "TRun"
'This is the starting point for all test runs. TRun.test will allow you to optional add tags and
' report formats while TRun.wip will call test with the @wip (work in progress) tag using the
' the verbose report format by default.
Option Explicit

Private m_step_implementations As Collection

Public Sub test(Optional filter_tag, Optional feature_filter, Optional report_format)
    
    Dim session As Senfgurke.TSession
    
    Set session = THelper.new_TSession
    session.run_test StepImplementations(session.ExecutionHooks), filter_tag, feature_filter, report_format, application_dir:=senfgurke_steps_workbook.path
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

Private Property Get StepImplementations(session_execution_hooks As Senfgurke.TExecutionHooks) As Collection
    
    Dim step_implementations As Variant
    Dim step_implementation_class As Variant

    Set m_step_implementations = New Collection
    'REGISTER all classes with STEP IMPLEMENTATIONS HERE >>>
    '-------------------------------------------------------
    step_implementations = Array(New Steps_cleanup_after_example, New Steps_collect_statistics, _
                                 New Steps_confirm_collection_member, New Steps_connect_steps_with_funct, _
                                 New Steps_Load_Feature_Files, _
                                 New Steps_parse_docstrings, _
                                 New Steps_Parse_Examples, _
                                 New Steps_Parse_Features, _
                                 New Steps_parse_rules, _
                                 New Steps_parse_step_expressions, _
                                 New Steps_parse_steps, _
                                 New Steps_parse_tables, New Steps_parse_outlines, _
                                 New Steps_parse_tags, _
                                 New Steps_predefined_steps, _
                                 New Steps_report, _
                                 New Steps_report_progress, _
                                 New Steps_report_statistics, _
                                 New Steps_report_verbose, New Steps_report_verbose_outlines, _
                                 New Steps_Run_Examples, New Steps_run_outline_example, _
                                 New Steps_Run_features, _
                                 New Steps_Run_Steps, _
                                 New Steps_run_tests, _
                                 New Steps_save_vars_in_context, _
                                 New Steps_show_step_template, _
                                 New Steps_support_functions, _
                                 New Steps_validate_expectations _
                                )
    Set m_step_implementations = New Collection
    For Each step_implementation_class In step_implementations
        On Error Resume Next
        'ignore error messages if the step implementationcalss hasn't declared an execution hook variable
        Set step_implementation_class.ExecutionHooks = session_execution_hooks
        On Error GoTo 0
        m_step_implementations.Add step_implementation_class
    Next
    Set StepImplementations = m_step_implementations
End Property
