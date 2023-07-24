VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_report_verbose_outlines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub When_the_results_from_the_execution_of_the_ou_84468CC2AB9B(example_context As TContext)
    'When the results from the execution of the outline examples are reported
    Dim session As TSession
    Dim gherkin_example As String
    
    gherkin_example = example_context.get_value("gherkin_outline_examples")
    'expect a session from the step stting the report format to verbose
    Set session = example_context.get_value("session")
    dry_run_outline_example gherkin_example, session, example_context
End Sub

Public Sub When_the_results_from_the_execution_of_the_ou_170DEFF10F3B(example_context As TContext, step_expressions As Collection)
    'When the results from the execution of the outline examples are reported with status {string}
    Dim session As TSession
    Dim gherkin_example As String
    
    gherkin_example = example_context.get_value("gherkin_outline_examples")
    'expect a session from the step stting the report format to verbose
    Set session = example_context.get_value("session")
    'default status for dry runs is OK
    dry_run_outline_example gherkin_example, session, example_context
End Sub

Private Sub dry_run_outline_example(gherkin_example As String, session As TSession, example_context As TContext)
    Dim example_runner As TExampleRunner
    Dim parsed_example As TExample
    Dim event_listener As EventListener
    
    Set parsed_example = TFeatureParser.parse_example(gherkin_example, THelper.new_TFeature, New Collection)
    'listen to all events logged for the given session
    Set event_listener = New EventListener
    Set event_listener.SenfgurkeEventLogger = session.EventLogger
    'run the outline examples without executing the steps (dry run)
    Set example_runner = THelper.new_TExampleRunner
    example_runner.setup session
    example_runner.run_outline_example parsed_example, True
End Sub

Public Sub Then_the_report_output_of_the_example_starts__64BB2A944AE4(example_context As TContext, step_expressions As Collection)
    'Then the report output of the example starts with {string}
    Dim session As TSession
    
    Set session = example_context.get_value("session")
    TSpec.expect(ExtraVBA.align_textblock(session.reporter.ConsoleOutput)).starts_with_text _
                    vbLf & step_expressions(1) & vbLf
End Sub
