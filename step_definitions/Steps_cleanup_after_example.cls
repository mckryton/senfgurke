VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_cleanup_after_example"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents StepStatusListener As TEventLogger
Attribute StepStatusListener.VB_VarHelpID = -1
Private FailedSteps As Collection
Private WithEvents TestSessionExecutionHooks As TExecutionHooks
Attribute TestSessionExecutionHooks.VB_VarHelpID = -1
Private AfterExampleCallsCount As Integer
Private ExampleExecutionStatus As String

Public Sub Given_an_example_that_should_not_run_into_fai_B68B15500956(example_context As TContext, step_expressions As Collection)
    'Given an example that should not run into failures {string}
    Dim sample_feature_gherkin As String
    Dim parsed_sample_feature As TFeature
    Dim parsed_example As TExample
    
    sample_feature_gherkin = "Feature: sample" & vbLf & _
                             "  Scenario: passing sample" & vbLf _
                             & step_expressions(1)
    Set parsed_sample_feature = Senfgurke.TFeatureParser.parse_feature(sample_feature_gherkin)
    Set parsed_example = parsed_sample_feature.sections(1)
    example_context.set_value parsed_example, "parsed_example"
End Sub

Public Sub Given_a_scenario_STR_6A3541DCE055(example_context As TContext, step_expressions As Collection)
    'Given a scenario {string}
    Dim sample_scenario_gherkin As String
        
    sample_scenario_gherkin = step_expressions(1)
    example_context.set_value sample_scenario_gherkin, "sample_scenario_gherkin"
End Sub

Public Sub When_the_result_for_running_every_step_of_the_4D22A6E0BB19(example_context As TContext, step_expressions As Collection)
    'When the result for running every step of the example is {string}
    Dim test_session As TSession
    Dim example_runner As TExampleRunner
    Dim parsed_example As TExample
    
    'create and remember a new test session for the scenario
    Set test_session = THelper.new_TSession
    example_context.set_value test_session, "test_session"
    'reset the counter for after examplehook calls
    AfterExampleCallsCount = 0
    'connect this class to the execution hooks from test session
    Set TestSessionExecutionHooks = test_session.ExecutionHooks
    'listen for run_step_finished events from test_session
    Set StepStatusListener = test_session.EventLogger
    Set FailedSteps = New Collection
    test_session.enable_silent_mode
    'execute the steps described in the scenario
    Set example_runner = THelper.new_TExampleRunner
    example_runner.setup test_session
    Set parsed_example = example_context.get_value("parsed_example")
    
    'executing the scenario with dry_run = true results with all steps returning OK
    example_runner.run_example parsed_example, dry_run:=True
    
    'expect that none of the steps described by the scenario has failed
    TSpec.expect(FailedSteps.Count).to_be 0
    Set FailedSteps = Nothing
End Sub

Public Sub When_the_scenario_succeeds_successfully_for_a_F7618495DC78(example_context As TContext, step_expressions As Collection)
    'When the scenario succeeds successfully for all steps {string}
        Dim sample_scenario_gherkin As String
    Dim test_session As TSession
    Dim example_runner As TExampleRunner
    
    'reset the test results
    AfterExampleCallsCount = 0
    ExampleExecutionStatus = vbNullString
    'retrieve the sample gherkin
    sample_scenario_gherkin = example_context.get_value("sample_scenario_gherkin")
    'create a separate session for running the sample scenario
    Set test_session = THelper.new_TSession
    example_context.set_value test_session, "test_session"
    'connect this class to the execution hooks from test session
    Set TestSessionExecutionHooks = test_session.ExecutionHooks
    'listen for run_step_finished events from test_session and save the step results in the FailedSteps collection
    Set StepStatusListener = test_session.EventLogger
    Set FailedSteps = New Collection
    'disable reporting on the console for the test session
    test_session.enable_silent_mode
    'execute the steps described in the scenario
    run_example sample_scenario_gherkin, test_session, dry_run:=True
    'verify that the given example succeeded successfully
    TSpec.expect(FailedSteps.Count).to_be 0
    Set FailedSteps = Nothing
End Sub

Public Sub When_the_scenario_fails_because_of_missing_st_416C361CBD50(example_context As TContext, step_expressions As Collection)
    'When the scenario fails because of missing step implementations {string}
    Dim sample_scenario_gherkin As String
    Dim test_session As TSession
    Dim example_runner As TExampleRunner
    
    'reset the test results
    AfterExampleCallsCount = 0
    ExampleExecutionStatus = vbNullString
    'retrieve the sample gherkin
    sample_scenario_gherkin = example_context.get_value("sample_scenario_gherkin")
    'create a separate session for running the sample scenario
    Set test_session = THelper.new_TSession
    example_context.set_value test_session, "test_session"
    'connect this class to the execution hooks from test session
    Set TestSessionExecutionHooks = test_session.ExecutionHooks
    'listen for run_step_finished events from test_session and save the step results in the FailedSteps collection
    Set StepStatusListener = test_session.EventLogger
    Set FailedSteps = New Collection
    'disable reporting on the console for the test session
    test_session.enable_silent_mode
    'execute the steps described in the scenario
    run_example sample_scenario_gherkin, test_session
    'verify that the given example really failed
    TSpec.expect(FailedSteps.Count).not_to_be 0
    Set FailedSteps = Nothing
End Sub

Private Sub run_example(sample_scenario_gherkin As String, test_session As Senfgurke.TSession, Optional dry_run As Boolean)
    Dim sample_feature_gherkin As String
    Dim parsed_sample_feature As TFeature
    Dim parsed_example As TExample
    Dim example_runner As TExampleRunner
    
    If IsMissing(dry_run) Then dry_run = False
    'put the example into a feature and parse it
    sample_feature_gherkin = "Feature: sample" & vbLf & _
                             "  Scenario: passing sample" & vbLf _
                             & sample_scenario_gherkin
    Set parsed_sample_feature = Senfgurke.TFeatureParser.parse_feature(sample_feature_gherkin)
    Set parsed_example = parsed_sample_feature.sections(1)
    'listen for run_step_finished events from test_session and save the step results in the FailedSteps collection
    Set StepStatusListener = test_session.EventLogger
    Set FailedSteps = New Collection
    'disable reporting on the console for the test session
    test_session.enable_silent_mode
    'execute the steps described in the scenario
    Set example_runner = THelper.new_TExampleRunner
    example_runner.setup test_session
    'run the example as is (no dry run)to trigger the expected events and hook functions
    If dry_run Then
        'use the dry run parameter to execute all steps with OK
        example_runner.run_example parsed_example, dry_run:=True
    Else
        example_runner.run_example parsed_example
    End If
End Sub

Public Sub Then_the_execution_status_of_the_scenario_is__F7D1AF58831B(example_context As TContext, step_expressions As Collection)
    'And the execution status of the scenario is attached as {string}
    TSpec.expect(ExampleExecutionStatus).to_be step_expressions(1)
End Sub

Public Sub Then_the_after_scenario_hook_function_is_call_FACA40D31E1C(example_context As TContext)
    'Then the after scenario hook function is called after the example
    TSpec.expect(AfterExampleCallsCount).not_to_be 0
End Sub

Private Sub Class_Initialize()
    Set FailedSteps = New Collection
    AfterExampleCallsCount = 0
End Sub

Private Sub Class_Terminate()
    Set FailedSteps = Nothing
End Sub

Private Sub TestSessionExecutionHooks_AfterExampleHook(log_event As Senfgurke.TEvent)
    AfterExampleCallsCount = AfterExampleCallsCount + 1
    ExampleExecutionStatus = log_event.SectionStatus
End Sub

Private Sub StepStatusListener_EventLogged(session_event As Senfgurke.TEvent)
    If session_event.EventName = "run_step_finished" And session_event.SectionStatus <> "OK" Then
        If session_event.SectionStatus <> "OK" Then
            FailedSteps.Add session_event.SectionStatus & " " & session_event.SectionName
        End If
    End If
End Sub
