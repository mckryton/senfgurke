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

Public Sub Given_a_feature_starting_with_random_text_in_STR_C4F336C986E5(step_expressions As Collection)
    'Given a feature starting with random text in {string}
    Dim loaded_feature As Collection
    
    TStepVars.feature_origin = step_expressions(1)
    Set loaded_feature = New Collection
    loaded_feature.Add "some random text", "feature_text"
    loaded_feature.Add TStepVars.feature_origin, "origin"
    Set TStepVars.loaded_features = New Collection
    TStepVars.loaded_features.Add loaded_feature
End Sub

Public Sub Given_a_feature_tagged_with_STR_with_one_example_C4F3ED80DA42(step_expressions As Collection)
    'Given a feature tagged with {string} with one example
    Dim loaded_feature As Collection
    
    Set loaded_feature = New Collection
    loaded_feature.Add step_expressions(1) & vbLf _
            & "Feature: tagged feature" & vbLf _
            & "  Example: sample" & vbLf _
            & "    Given a valid step", "feature_text"
    Set TStepVars.loaded_features = New Collection
    TStepVars.loaded_features.Add loaded_feature
End Sub

Public Sub Given_an_untagged_feature_with_one_example_E3F4277DD399()
    'And an un-tagged feature with one example
    Dim loaded_feature As Collection
    
    Set loaded_feature = New Collection
    loaded_feature.Add "Feature: un-tagged feature" & vbLf _
            & "  Example: sample" & vbLf _
            & "    Given a valid step", "feature_text"
    Set TStepVars.loaded_features = New Collection
    TStepVars.loaded_features.Add loaded_feature
End Sub

Public Sub When_Senfgurke_executes_the_feature_7D772B9CDDDE()
    'When Senfgurke executes the feature
    Set TStepVars.Session = New TSession
    TStepVars.Session.Reporter.enable_silent_mode
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

Public Sub Then_only_the_tagged_feature_is_executed_F87C2CEF089C()
    'Then only the tagged feature is executed
    Dim feature_events As Collection
    Dim log_event As Variant
    
    Set feature_events = New Collection
    For Each log_event In TStepVars.Session.statistics.event_log
        If log_event("type") = LOG_TYPE_FEATURE Then feature_events.Add log_event
    Next
    TSpec.expect(feature_events.Count).to_be 1
    TSpec.expect(feature_events(1)("msg")).includes_text "tagged feature"
End Sub
