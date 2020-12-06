Attribute VB_Name = "TConfig"
Option Explicit

Dim m_step_implementations As Collection

Public Property Get StepImplementations() As Collection
    
    Dim step_implementation_class As Variant
    
    If m_step_implementations Is Nothing Then
        Set m_step_implementations = New Collection
        'add all classes with step implementations here:
        For Each step_implementation_class In Array(New Steps_Run_Examples, New Steps_Parse_Features, _
                        New Steps_Load_Feature_Files, New Steps_report_verbose, New Steps_Run_Steps, _
                        New Steps_Parse_Examples, New Steps_Run_features, New Steps_confirm_collection_member, _
                        New Steps_make_steps_executable)
            m_step_implementations.Add step_implementation_class
        Next
    End If
    Set StepImplementations = m_step_implementations
End Property

Public Property Get MaxStepFunctionNameLength() As Variant
    
    'VBA compiler crashes under MacOS (v16.41) if a function name is longer than 63 characters
    ' specified max value is 255 characters
    MaxStepFunctionNameLength = 63
End Property
