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

Public Sub Given_a_step_And_another_step_3B012F56407C()
    'Given a step "And another step"
    Set example_step = New TStep
    example_step.Head = "And"
    example_step.Name = "another step"
End Sub

Public Sub Given_the_type_of_the_previous_step_is_Given_F33FC2478AD0()
    'Given the type of the previous step is "Given"
    example_step.SType = "Given"
End Sub

Public Sub Given_a_step_But_no_sound_occurs_EB3A6373A9FC()
    'Given a step "But no sound occurs"
    Set example_step = New TStep
    example_step.Head = "But"
    example_step.Name = "no sound occurs"
End Sub

Public Sub Given_the_type_of_the_previous_step_is_When_F33FC2476EF3()
    'Given the type of the previous step is "When"
    example_step.SType = "When"
End Sub

Public Sub Given_a_step_STR_F5C6F4BA1D15(step_parameters As Collection)
    'Given a step "Given the value of pi is 3.14"
    Set example_step = New TStep
    With example_step
        .Head = "Given"
        .SType = "Given"
        .Name = Right(step_parameters(1), Len(step_parameters(1)) - Len("Given "))
    End With
End Sub

Public Sub When_the_function_name_for_the_second_step_is_calc_7916E785717F()
    'When the function name for the second step is calculated
    m_step_function_name = example_step.get_step_function_name
End Sub

Public Sub When_the_function_for_the_step_is_calculated_41249A4C9E7F()
    'When the function for the step is calculated
    m_step_function_name = example_step.get_step_function_name
End Sub

Public Sub Then_the_function_name_starts_with_STR_4C16558C4B9D(step_parameters As Collection)
    'Then the function name starts with "Given_the_name_of_the_first_day_of_the_week_is_STR"
    TSpec.expect(Left(m_step_function_name, Len(step_parameters(1)))).to_be step_parameters(1)
End Sub
