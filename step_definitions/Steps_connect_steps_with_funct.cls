VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_connect_steps_with_funct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_a_list_of_steps_FF6B22C2A87B(example_context As TContext, data_table As TDataTable)
    'Given a list of steps
    Dim step As TStep
    Dim step_list As Collection
    Dim table_row As Variant
    
    Set step_list = New Collection
    For Each table_row In data_table.table_rows
        Set step = TStepParser.parse_step_line(CStr(table_row("example_step")), step_list)
        step_list.Add step
    Next
    example_context.set_value step_list, "example_steps"
End Sub

Public Sub When_the_function_name_for_the_last_step_is_calcul_3D1B9F9C0F3D(example_context As TContext)
    'When the function name for the last step is calculated
    Dim Feature As TFeature
    Dim Example As TExample
    Dim last_step As TStep
    
    example_context.set_value "Feature: sample" & vbLf & example_context.get_value("gherkin_text"), "gherkin_text"
    Set Feature = TFeatureParser.parse_feature(example_context.get_value("gherkin_text"))
    Set Example = Feature.sections(1)
    Set last_step = Example.steps(Example.steps.Count)
    example_context.set_value last_step.get_step_function_name, "step_function_name"
End Sub

Public Sub When_the_function_for_the_step_is_calculated_41249A4C9E7F(example_context As TContext)
    'When the function for the step is calculated
    example_context.set_value example_context.get_value("example_step").get_step_function_name, "step_function_name"
End Sub

Public Sub When_the_a_matching_function_for_those_steps_is_re_AB345B8299F4(example_context As TContext)
    'When the a matching function for those steps is requested
    Dim function_names As Collection
    Dim step_list As Collection
    Dim step As Variant
    
    Set step_list = example_context.get_value("example_steps")
    Set function_names = New Collection
    For Each step In step_list
        function_names.Add step.get_step_function_name
    Next
    example_context.set_value function_names, "function_names"
End Sub

Public Sub Then_the_function_name_starts_with_STR_4C16558C4B9D(example_context As TContext, step_expressions As Collection)
    'Then the function name starts with "Given_the_name_of_the_first_day_of_the_week_is_STR"
    TSpec.expect(Left(example_context.get_value("step_function_name"), Len(step_expressions(1)))).to_be step_expressions(1)
End Sub

Public Sub Then_the_matching_function_name_is_STR_45E90A644CD7(example_context As TContext, step_expressions As Collection)
    'Then the matching function name is {string}
    TSpec.expect(example_context.get_value("step_function_name")).to_be step_expressions(1)
End Sub

Public Sub Then_the_same_function_name_is_returned_for_all_st_2DD39A9E8F72(example_context As TContext)
    'Then the same function name is returned for all steps
    Dim function_names As Collection
    Dim reference_name As String
    Dim function_name As Variant
    Dim index As Integer
    
    Set function_names = example_context.get_value("function_names")
    reference_name = function_names(1)
    For index = 2 To function_names.Count
        TSpec.expect(function_names(index)).to_be reference_name
    Next
End Sub
