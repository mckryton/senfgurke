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

Public Sub When_Senfgurke_executes_the_feature_7D772B9CDDDE()
    'When Senfgurke executes the feature
    Set TStepVars.Session = New TSession
    TStepVars.Session.Reporter.enable_silent_mode
    TStepVars.Session.run_test features_as_text:=TStepVars.loaded_features
End Sub

Public Sub Then_the_error_STR_is_reported_80FDB77BCB22(step_expressions As Collection)
    'Then the error {string} is reported
    TSpec.expect(TStepVars.Session.Reporter.ConsoleOutput).includes_text step_expressions(1)
End Sub

Public Sub Then_the_name_of_the_feature_file_is_reported_as_l_1A076B1EB8D5()
    'And the name of the feature file is reported as location of the error
    TSpec.expect(TStepVars.Session.Reporter.ConsoleOutput).includes_text TStepVars.feature_origin
End Sub
