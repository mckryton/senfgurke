VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Run_Steps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_step_result As Variant
Dim m_step_function_template As String
Dim m_step_function_name As String

Public Sub Given_a_step_returning_PENDING_without_a_pending_m_F201CAEBCFB1()
    'Given a step returning PENDING without a pending message
    Set TStepVars.example_step = New TStep
    TStepVars.example_step.SType = STEP_TYPE_GIVEN
    TStepVars.example_step.Name = "a pending step"
    'won't create code dynamically, so step "Given a pending step" is hard coded below
End Sub

Public Sub Given_a_pending_step_6A3533BBBFE5()
    pending
End Sub

Public Sub Given_an_empty_step_8A7408D743D5()

End Sub

Public Sub Given_a_valid_step_8A74152FD2F9()
    TSpec.expect(1).to_be 1
End Sub

Public Sub Given_an_invalid_step_CF4B201BCF1B()
    TSpec.expect(1).to_be 2
End Sub

Public Sub Given_a_step_with_an_empty_implementation_F776BB776C61()
    'Given a step with an empty implementation
    Set TStepVars.example_step = New TStep
    TStepVars.example_step.SType = STEP_TYPE_GIVEN
    TStepVars.example_step.Name = "an empty step"
    'the step function itself can't be created dynamically
End Sub

Public Sub Given_a_step_with_a_working_implementation_C0368FD3CAAA()
    'Given a step with a working implementation
    Set TStepVars.example_step = New TStep
    TStepVars.example_step.SType = STEP_TYPE_GIVEN
    TStepVars.example_step.Name = "a valid step"
    'the step function itself can't be created dynamically
End Sub

Public Sub Given_a_step_with_a_failing_implementation_C036A44ACAAA()
    'Given a step with a failing implementation
    Set TStepVars.example_step = New TStep
    TStepVars.example_step.SType = STEP_TYPE_GIVEN
    TStepVars.example_step.Name = "an invalid step"
    'the step function itself can't be created dynamically
End Sub

Public Sub Given_a_step_without_an_implementation_F76ADB85CAAA()
    'Given a step without an implementation
    Set TStepVars.example_step = New TStep
    TStepVars.example_step.SType = STEP_TYPE_GIVEN
    TStepVars.example_step.Name = "a random step name " & CStr(Now())
End Sub

Public Sub Given_a_step_Given_a_valid_step_EB3A0AE3CD52()
    'Given a step "Given a valid step"
    Set TStepVars.example_step = New TStep
    TStepVars.example_step.SType = STEP_TYPE_GIVEN
    TStepVars.example_step.Name = "a valid step"
End Sub

Public Sub Given_a_step_Given_a_step_812EA637D130()
    'Given a step "Given a step"
    Set TStepVars.example_step = New TStep
    TStepVars.example_step.SType = STEP_TYPE_GIVEN
    TStepVars.example_step.Name = "a step"
End Sub

Public Sub When_the_step_is_executed_CEB06B444334()
    'When the step is executed
    m_step_result = TStepRunner.run_step(TStepVars.example_step)
End Sub

Public Sub When_the_code_template_for_the_step_implementation_0EC5AFD47941()
    'When the code template for the step implementation is requested
    m_step_function_template = TStepVars.example_step.get_step_function_template
End Sub

Public Sub When_the_step_is_translated_into_a_function_name_0F03C2292FFD()
    'When the step is translated into a function name
    m_step_function_name = TStepVars.example_step.get_step_function_name
End Sub

Public Sub Then_the_execution_result_is_STR_80FDCF952A5A(step_expressions As Collection)
    'Then the execution result is {string}
    TSpec.expect(m_step_result(0)).to_be step_expressions(1)
End Sub

Public Sub Then_the_code_template_for_the_step_implementation_46484D426FED()
    'Then the code template for the step implementation is "Public Sub Given_a_valid_step_8A74152FD2F9()" & vblf & "    'Given a valid step" & vblf & "" & vblf & "End Sub"
    TSpec.expect(m_step_function_template).to_be _
        "Public Sub Given_a_valid_step_8A74152FD2F9()" & vbLf _
        & "    'Given a valid step" & vbLf _
        & vbLf _
        & "End Sub"
End Sub

Public Sub Then_the_name_of_the_resulting_function_is_Given_a_E0B741C87B98()
    'Then the name of the resulting function is "Given_a_step_C7224E350E70"
    TSpec.expect(m_step_function_name).to_be "Given_a_step_C7224E350E70"
End Sub

