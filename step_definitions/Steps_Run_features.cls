VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Run_features"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_a_feature_has_a_background_step_F0A675096BE8()
    'Given a feature has a background step
    TStepVars.gherkin_text = "Feature: sample feature" & vbLf _
                                & vbLf _
                                & "  Background:" & vbLf _
                                & "    Given a background step" & vbLf
End Sub

Public Sub Given_the_feature_has_two_examples_with_on_step_DE25FF3D87A5()
    'And the feature has two examples with on step
    TStepVars.gherkin_text = TStepVars.gherkin_text _
                                & vbLf _
                                & "  Example: example one" & vbLf _
                                & "    Given a step" & vbLf _
                                & vbLf _
                                & "  Example: example two" & vbLf _
                                & "    Given a step"
End Sub

Public Sub Given_a_step_C72276450E70()
    'Given a step
    
End Sub

Public Sub Given_a_background_step_FF6B02C4DC47()
    'Given a background step
    
End Sub

Public Sub When_the_feature_is_executed_8C8F454A90BC()
    'When the feature is executed
    
    Dim parsed_feature As TFeature
    Dim feature_runner As TFeatureRunner
    
    Set feature_runner = New TFeatureRunner
    Set TStepVars.Session = New TSession
    TStepVars.Session.Reporter.enable_silent_mode
    feature_runner.setup TStepVars.Session
    Set parsed_feature = TFeatureParser.parse_feature(TStepVars.gherkin_text)
    Set TStepVars.feature_statistics = feature_runner.run_feature(parsed_feature)
End Sub

Public Sub Then_the_background_step_is_executed_once_before_e_AD9FFABBC1CE()
    'Then the background step is executed once before each example
    TSpec.expect(TStepVars.feature_statistics(1)).contains_member Array("OK", "Given a background step")
    TSpec.expect(TStepVars.feature_statistics(2)).contains_member Array("OK", "Given a background step")
End Sub
