Attribute VB_Name = "TConfig"
Option Explicit

Dim m_step_implementations As Collection

Public Property Get StepImplementations() As Collection
    
    Dim step_implementation_class As Variant
    
    If m_step_implementations Is Nothing Then
        Set m_step_implementations = New Collection
        'add all classes with step implementations here:
        For Each step_implementation_class In Array(New Steps_run_tests, New Steps_Run_Examples, New Steps_Run_Steps, _
                        New Steps_Run_features, New Steps_Load_Feature_Files, New Steps_report_verbose, _
                        New Steps_Parse_Features, New Steps_Parse_Examples, New Steps_validate_expectations, _
                        New Steps_make_steps_executable, New Steps_parse_steps, New Steps_parse_step_expressions, _
                        New Steps_parse_docstrings, New Steps_parse_tags, New Steps_parse_rules, _
                        New Steps_report, New Steps_report_progress, New Steps_report_statistics, _
                        New Steps_collect_statistics, New Steps_support_functions)
            m_step_implementations.Add step_implementation_class
        Next
    End If
    Set StepImplementations = m_step_implementations
End Property

Public Property Let StepImplementations(new_stepimplementations As Collection)
    Set m_step_implementations = new_stepimplementations
End Property

Public Property Get MaxStepFunctionNameLength() As Variant
    
    'VBA compiler crashes under MacOS (v16.41) if a function name is longer than 63 characters
    ' specified max value is 255 characters
    MaxStepFunctionNameLength = 63
End Property
