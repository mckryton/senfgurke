VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_run_tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_a_new_test_run_6A3558F7A1DD()
    'Given a new test run
    Set TStepVars.loaded_features = Nothing
End Sub

Public Sub Given_a_feature_was_loaded_as_STR_A7346B4E7CD1(step_expressions As Collection)
    'Given a feature was loaded as {string}
    If TStepVars.loaded_features Is Nothing Then Set TStepVars.loaded_features = New Collection
    TSupport.load_feature_from_text step_expressions(1), TStepVars.loaded_features
End Sub

Public Sub Given_a_feature_was_loaded_from_a_file_STR_STR_725154F9143A(step_expressions As Collection)
    'Given a feature was loaded from a file {string} {string}
    TStepVars.feature_origin = step_expressions(1)
    If TStepVars.loaded_features Is Nothing Then Set TStepVars.loaded_features = New Collection
    TSupport.load_feature_from_text step_expressions(2), TStepVars.loaded_features, TStepVars.feature_origin
End Sub

Public Sub When_Senfgurke_executes_the_feature_7D772B9CDDDE()
    'When Senfgurke executes the feature
    Set TStepVars.Session = New TSession
    TStepVars.Session.Reporter.enable_silent_mode
    TSpec.expect(TStepVars.loaded_features).not_to_be_nothing
    TStepVars.Session.run_test features_as_text:=TStepVars.loaded_features
End Sub

Public Sub When_a_test_is_started_with_STR_as_parameter_EA8A08D9425A(step_expressions As Collection)
    'When a test is started with {string} as parameter
    Set TStepVars.Session = New TSession
    TStepVars.Session.Reporter.enable_silent_mode
    TStepVars.Session.statistics.enable_event_log
    TStepVars.Session.run_test features_as_text:=TStepVars.loaded_features, filter_tag:=step_expressions(1)
End Sub

Public Sub Then_the_error_STR_is_reported_80FDB77BCB22(step_expressions As Collection)
    'Then the error {string} is reported
    TSpec.expect(TStepVars.Session.Reporter.ConsoleOutput).includes_text step_expressions(1)
End Sub

Public Sub Then_the_name_of_the_feature_file_is_reported_as_l_1A076B1EB8D5()
    'And the name of the feature file is reported as location of the error
    TSpec.expect(TStepVars.Session.Reporter.ConsoleOutput).includes_text TStepVars.feature_origin
End Sub

Public Sub Then_only_the_example_from_the_tagged_feature_was__858E135D9F78()
    'Then only the example from the tagged feature was executed
    Dim example_events As Collection
    
    Set example_events = TStepVars.Session.statistics.filter_events(LOG_TYPE_EXAMPLE)
    TSpec.expect(example_events.Count).to_be 1
    TSpec.expect(example_events(1)("msg")).includes_text "sample from tagged feature"
End Sub

Public Sub Then_only_the_example_from_the_tagged_rule_was_exe_0F822B53DF54()
    'Then only the example from the tagged rule was executed
    Dim example_events As Collection
    
    Set example_events = TStepVars.Session.statistics.filter_events(LOG_TYPE_EXAMPLE)
    TSpec.expect(example_events.Count).to_be 1
    TSpec.expect(example_events(1)("msg")).includes_text "example assigned to the rule"
End Sub

Public Sub Then_no_example_was_executed_16BF2F2090BC()
    'Then no example was executed
    Dim example_events As Collection
    
    Set example_events = TStepVars.Session.statistics.filter_events(LOG_TYPE_EXAMPLE)
    TSpec.expect(example_events.Count).to_be 0
End Sub

