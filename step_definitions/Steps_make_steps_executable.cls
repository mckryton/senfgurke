VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_make_steps_executable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub When_the_function_name_for_the_last_step_is_calcul_3D1B9F9C0F3D()
    'When the function name for the last step is calculated
    Dim feature As TFeature
    Dim example As TExample
    Dim last_step As TStep
    
    TStepVars.gherkin_text = "Feature: sample" & vbLf & TStepVars.gherkin_text
    Set feature = TFeatureParser.parse_feature(TStepVars.gherkin_text)
    Set example = feature.sections(1)
    Set last_step = example.Steps(example.Steps.Count)
    m_step_function_name = last_step.get_step_function_name
End Sub

Public Sub When_the_function_for_the_step_is_calculated_41249A4C9E7F()
    'When the function for the step is calculated
    m_step_function_name = TStepVars.example_step.get_step_function_name
End Sub

Public Sub Then_the_function_name_starts_with_STR_4C16558C4B9D(step_parameters As Collection)
    'Then the function name starts with "Given_the_name_of_the_first_day_of_the_week_is_STR"
    TSpec.expect(Left(m_step_function_name, Len(step_parameters(1)))).to_be step_parameters(1)
End Sub

