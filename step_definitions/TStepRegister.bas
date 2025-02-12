Attribute VB_Name = "TStepRegister"
Option Explicit


'REGISTER all classes with STEP IMPLEMENTATIONS HERE >>>
Public Function get_step_definition_classes() As Variant
    get_step_definition_classes = Array(New Steps_cleanup_after_example, New Steps_collect_statistics, _
                                         New Steps_assure_collection_members, New Steps_connect_steps_with_funct, _
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
                                         New Steps_assure_expectations, New Steps_custom_err_msg_expectatio _
                                        )
End Function
