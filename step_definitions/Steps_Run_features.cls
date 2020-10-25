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

Public Sub Given_a_feature_has_a_background_step_FD3D5509A1FD()
    'Given a feature has a background step
    TStepVars.feature_text = "Feature: sample feature" & vbLf _
                                & vbLf _
                                & "  Background:" & vbLf _
                                & "    Given a background step" & vbLf
End Sub

Public Sub Given_the_feature_has_two_examples_with_on_step_0A1C38A7165C()
    'Given the feature has two examples with on step
    TStepVars.feature_text = TStepVars.feature_text _
                                & vbLf _
                                & "  Example: example one" & vbLf _
                                & "    Given a step" & vbLf _
                                & vbLf _
                                & "  Example: example two" & vbLf _
                                & "    Given a step"
End Sub

Public Sub Given_a_step_C7224E350E70()
    'Given a step
End Sub

Public Sub Given_a_background_step_CF4BC54A1652()
    'Given a background step
End Sub

Public Sub When_the_feature_is_executed_A6DB756AA298()
    'When the feature is executed
    
    Dim parsed_feature As TFeature
    
    Set parsed_feature = TFeatureParser.parse_feature(TStepVars.feature_text)
    Set TStepVars.feature_statistics = TFeatureRunner.run_feature(parsed_feature, silent:=True)
End Sub

Public Sub Then_the_background_step_is_executed_once_before_e_36EB01695F3D()
    'Then the background step is executed once before each example
    TSpec.expect(TStepVars.feature_statistics(1)).contains Array("OK", "Given a background step")
    TSpec.expect(TStepVars.feature_statistics(2)).contains Array("OK", "Given a background step")
End Sub
