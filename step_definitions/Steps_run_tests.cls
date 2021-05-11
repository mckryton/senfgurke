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

Public Sub Given_a_feature_was_loaded_as_STR_A7346B4E7CD1(context As TContext, step_expressions As Collection)
    'Given a feature was loaded as {string}
    If Not context.value_exists("loaded_features") Then
        context.set_value New Collection, "loaded_features"
    End If
    TSupport.load_feature_from_text step_expressions(1), context.get_value("loaded_features")
End Sub

Public Sub Given_a_feature_was_loaded_from_a_file_STR_STR_725154F9143A(context As TContext, step_expressions As Collection)
    'Given a feature was loaded from a file {string} {string}
    context.set_value step_expressions(1), "feature_origin"
    If Not context.value_exists("loaded_features") Then
        context.set_value New Collection, "loaded_features"
    End If
    TSupport.load_feature_from_text step_expressions(2), context.get_value("loaded_features"), context.get_value("feature_origin")
End Sub

Public Sub When_Senfgurke_executes_the_feature_7D772B9CDDDE(context As TContext)
    'When Senfgurke executes the feature
    Set TStepVars.session = New TSession
    TStepVars.session.Reporter.enable_silent_mode
    TStepVars.session.run_test features_as_text:=context.get_value("loaded_features")
End Sub

Public Sub When_a_test_is_started_with_STR_as_parameter_EA8A08D9425A(context As TContext, step_expressions As Collection)
    'When a test is started with {string} as parameter
    Set TStepVars.session = New TSession
    TStepVars.session.Reporter.enable_silent_mode
    TStepVars.session.statistics.enable_event_log
    TStepVars.session.run_test features_as_text:=context.get_value("loaded_features"), filter_tag:=step_expressions(1)
End Sub

Public Sub Then_the_error_STR_is_reported_80FDB77BCB22(context As TContext, step_expressions As Collection)
    'Then the error {string} is reported
    TSpec.expect(TStepVars.session.Reporter.ConsoleOutput).includes_text step_expressions(1)
End Sub

Public Sub Then_the_name_of_the_feature_file_is_reported_as_l_1A076B1EB8D5(context As TContext)
    'And the name of the feature file is reported as location of the error
    TSpec.expect(TStepVars.session.Reporter.ConsoleOutput).includes_text context.get_value("feature_origin")
End Sub

Public Sub Then_only_the_example_from_the_tagged_feature_was__858E135D9F78(context As TContext)
    'Then only the example from the tagged feature was executed
    Dim example_events As Collection
    
    Set example_events = TStepVars.session.statistics.filter_events(LOG_TYPE_EXAMPLE)
    TSpec.expect(example_events.Count).to_be 1
    TSpec.expect(example_events(1)("msg")).includes_text "sample from tagged feature"
End Sub

Public Sub Then_only_the_example_from_the_tagged_rule_was_exe_0F822B53DF54(context As TContext)
    'Then only the example from the tagged rule was executed
    Dim example_events As Collection
    
    Set example_events = TStepVars.session.statistics.filter_events(LOG_TYPE_EXAMPLE)
    TSpec.expect(example_events.Count).to_be 1
    TSpec.expect(example_events(1)("msg")).includes_text "example assigned to the rule"
End Sub

Public Sub Then_no_example_was_executed_16BF2F2090BC(context As TContext)
    'Then no example was executed
    Dim example_events As Collection
    
    Set example_events = TStepVars.session.statistics.filter_events(LOG_TYPE_EXAMPLE)
    TSpec.expect(example_events.Count).to_be 0
End Sub

