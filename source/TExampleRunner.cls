VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TExampleRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_session As TSession
Private m_parent_runner As TFeatureRunner

Private Sub Class_Terminate()
    Set m_session = Nothing
    Set m_parent_runner = Nothing
End Sub

Public Sub setup(session As TSession, Optional parent_runner)
    Set m_session = session
    If IsMissing(parent_runner) Then
        'setup a new feature runner when example runner is used in a standalone setting (e.g. for testing)
        Set parent_runner = New TFeatureRunner
        parent_runner.setup session
    End If
    Set m_parent_runner = parent_runner
End Sub

Public Sub run_example(example As TExample, Optional outline_parameter_names, Optional outline_parameter_set, Optional dry_run)
    Dim step_result As Variant
    Dim err_msg As String
    Dim step As TStep
    Dim step_status As String
    Dim previous_step_status As String
    Dim step_runner As TStepRunner
    Dim context As TContext
    Dim original_step_line As String
    Dim example_status As String

    If IsMissing(dry_run) Then dry_run = False
    step_result = vbNullString
    Set step_runner = New TStepRunner
    step_runner.setup m_session, Me
    Set context = New TContext
    step_status = vbNullString
    example_status = vbNullString
    For Each step In example.steps
        err_msg = vbNullString
        previous_step_status = step_status
        If IsMissing(outline_parameter_names) Then
            'example doesn't have an outline - just run the step
            step_result = step_runner.run_step(step, context, dry_run)
        Else
            original_step_line = step.OriginalStepLine
            'change the original step definition according to the outline parameter(s)
            step.insert_outline_parameters outline_parameter_names, outline_parameter_set
            'repeat parsing the step name
            step.reset_name
            step_result = step_runner.run_step(step, context, dry_run)
            'reset the step definition back to step line including placeholders
            step.OriginalStepLine = original_step_line
        End If
        step_status = CStr(step_result(0))
        If previous_step_status <> STATUS_OK And previous_step_status <> vbNullString Then
            ' mark all steps as skipped when the last step wasn't OK
            If step_status <> STATUS_MISSING Then step_status = STATUS_SKIPPED
        End If
        If Not step_status = STATUS_OK And Not step_status = STATUS_SKIPPED Then
            'step failed -> get the failure message
            err_msg = step_result(1)
            'this is a hack because of the broken exception handling in vba (e.g. vba will overwrite the custom err descriptions)
            If step_status = STATUS_PENDING Then err_msg = TSpec.LastFailMsg
        End If
        log_event EVENT_RUN_STEP_FINISHED, err_msg, step.OriginalStepDefinition, SECTION_TYPE_STEP, step_status
    Next
    'consider the example to be OK only if the status of the last step was OK
    If step_status = STATUS_OK Then example_status = STATUS_OK Else example_status = STATUS_FAIL
    log_event EVENT_RUN_EXAMPLE_FINISHED, example.description, example.OriginalHeadline, SECTION_TYPE_EXAMPLE, example_status
End Sub

Public Sub run_outline_example(outline_example As TExample, Optional dry_run)
    Dim step As Variant
    Dim outline As Variant
    Dim outline_name As String
    Dim outline_table As TDataTable
    Dim outline_parameter_set As Variant
    Dim table_row As String
    
    'dry run is for testing the procedure when the result of the steps doesn't matter
    If IsMissing(dry_run) Then dry_run = False
    log_event EVENT_RUN_OUTLINE_EXAMPLE_STARTED, outline_example.description, outline_example.OriginalHeadline, SECTION_TYPE_EXAMPLE
    'inform about the steps to repeat
    For Each step In outline_example.steps
        log_event EVENT_RUN_OUTLINE_STEP_DECLARED, vbNullString, step.OriginalStepDefinition, SECTION_TYPE_STEP
    Next
    'every outline is represented by its own unique outline table
    For Each outline In outline_example.Outlines
        outline_name = Trim("Examples: " & outline("name"))
        log_event EVENT_RUN_OUTLINE_TABLE_STARTED, outline("table").get_table_header_as_text(), outline_name, SECTION_TYPE_OUTLINE
        Set outline_table = outline("table")
        For Each outline_parameter_set In outline_table.table_rows
            table_row = outline_table.get_table_row_as_text(outline_parameter_set)
            log_event EVENT_RUN_OUTLINE_ROW_STARTED, table_row, outline_name, SECTION_TYPE_OUTLINE
            run_example outline_example, outline_table.column_names, outline_parameter_set, dry_run
        Next
    Next
    log_event EVENT_RUN_OUTLINE_EXAMPLE_FINISHED, outline_example.description, outline_example.OriginalHeadline, SECTION_TYPE_EXAMPLE
End Sub

Public Sub log_event(event_name As String, Optional event_msg, Optional section_name, Optional section_type, _
                        Optional section_status, Optional feature_file, Optional log_time)
    m_parent_runner.log_event event_name, event_msg, section_name, section_type, section_status, feature_file, log_time
End Sub
