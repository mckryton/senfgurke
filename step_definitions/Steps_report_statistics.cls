VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_report_statistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_a_test_run_took_INT_ms_A7563698E569(example_context As TContext, step_expressions As Collection)
    'And a test run took {integer} ms
    example_context.set_value CLng(step_expressions(1)), "duration"
End Sub

Public Sub Given_one_example_in_this_run_had_INT_passed__800581B31566(example_context As TContext, step_expressions As Collection)
    'And one example in this run had {integer} passed, {integer} failed, {integer} missing and {integer} pending steps
    Dim index As Integer
    Dim session As TSession
   
    Set session = example_context.get_value("session")
    For index = 1 To step_expressions(1)
        session.log_event EVENT_RUN_STEP_FINISHED, vbNullString, "dummy step", SECTION_TYPE_STEP, STATUS_OK
    Next
    For index = 1 To step_expressions(2)
        session.log_event EVENT_RUN_STEP_FINISHED, vbNullString, "dummy step", SECTION_TYPE_STEP, STATUS_FAIL
    Next
    For index = 1 To step_expressions(3)
        session.log_event EVENT_RUN_STEP_FINISHED, vbNullString, "dummy step", SECTION_TYPE_STEP, STATUS_MISSING
    Next
    For index = 1 To step_expressions(4)
        session.log_event EVENT_RUN_STEP_FINISHED, vbNullString, "dummy step", SECTION_TYPE_STEP, STATUS_PENDING
    Next
End Sub

Public Sub When_the_statistics_are_reported_418C385A7E0B(example_context As TContext)
    'When the statistics are reported
    Dim session As TSession
    Dim statistic_summary As String
    Dim duration As Long
   
    Set session = example_context.get_value("session")
    duration = example_context.get_value("duration")
    statistic_summary = session.statistics.get_statistics_summary & vbLf & session.get_duration_as_text(duration)
    session.log_event EVENT_RUN_SESSION_FINISHED, statistic_summary
End Sub

Public Sub Then_the_resulting_report_output_ends_with_ST_422DE5AFAABA(example_context As TContext, step_expressions As Collection)
    'Then the resulting report output ends with {string}
    Dim console_output As String
    Dim session As TSession
   
    Set session = example_context.get_value("session")
    console_output = ExtraVBA.trim_linebreaks(session.reporter.ConsoleOutput)
    Replace session.reporter.ConsoleOutput, vbCr & vbLf, vbLf
    TSpec.expect(Right(console_output, Len(step_expressions(1)))).to_be step_expressions(1)
End Sub
