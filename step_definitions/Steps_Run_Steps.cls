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

Private m_step_result As Variant
Private m_step_function_template As String
Private m_step_function_name As String

Public Sub Given_a_pending_step_6A3533BBBFE5()
    pending
End Sub

Public Sub Given_an_empty_step_6A35DCCE89C0()
    'Given an empty step
    
End Sub

Public Sub Given_a_valid_step_8A74152FD2F9()
    TSpec.expect(1).to_be 1
End Sub

Public Sub Given_an_invalid_step_CF4B201BCF1B()
    TSpec.expect(1).to_be 2
End Sub

Public Sub Given_a_step_returning_PENDING_without_a_pending_m_F201CAEBCFB1()
    'Given a step returning PENDING without a pending message
    Set TStepVars.example_step = TFeatureParser.create_step("Given a pending step")
    'won't create code dynamically, so step "Given a pending step" is hard coded
End Sub

Public Sub Given_a_step_with_an_empty_implementation_BC7B46584C5E()
    'Given a step with an empty implementation
    Set TStepVars.example_step = TFeatureParser.create_step("Given an empty step")
    'the step function itself can't be created dynamically and has to be available
End Sub

Public Sub Given_a_step_with_a_working_implementation_05BC882CCAAA()
    'Given a step with a working implementation
    Set TStepVars.example_step = TFeatureParser.create_step("Given a valid step")
    'the step function itself can't be created dynamically and has to be available
End Sub

Public Sub Given_a_step_with_a_failing_implementation_05BCA3B5CAAA()
    'Given a step with a failing implementation
    Set TStepVars.example_step = TFeatureParser.create_step("Given an invalid step")
    'the step function itself can't be created dynamically and has to be available
End Sub

Public Sub Given_a_step_without_an_implementation_FAF18488CAAA()
    'Given a step without an implementation
    Set TStepVars.example_step = TFeatureParser.create_step("Given a random step name " & CStr(Now))
End Sub

Public Sub Given_a_step_STR_F5C6F4BA1D15(step_expressions As Collection)
    'Given a step {string}
    Set TStepVars.example_step = TFeatureParser.create_step(step_expressions(1))
End Sub

Public Sub When_the_step_is_executed_CEB06B444334()
    'When the step is executed
    Dim step_runner As TStepRunner
    
    Set step_runner = New TStepRunner
    Set TStepVars.Session = New TSession
    TStepVars.Session.Reporter.enable_silent_mode
    step_runner.setup TStepVars.Session
    m_step_result = step_runner.run_step(TStepVars.example_step)
End Sub

Public Sub When_the_code_template_for_the_step_implementation_7D39D6705DD8()
    'When the code template for the step implementation is requested
    m_step_function_template = TStepVars.example_step.get_step_function_template
End Sub

Public Sub When_the_step_is_translated_into_a_function_name_D820185C4AF6()
    'When the step is translated into a function name
    m_step_function_name = TStepVars.example_step.get_step_function_name
End Sub


Public Sub Then_the_execution_result_is_STR_80FDCF952A5A(step_expressions As Collection)
    'Then the execution result is {string}
    TSpec.expect(m_step_result(0)).to_be step_expressions(1)
End Sub

Public Sub Then_the_code_template_for_the_step_implementation_9C79C155DB8E(step_expressions As Collection)
    'Then the code template for the step implementation is {string}
    'trim template lines, because docstring parameter is normalized too
    Dim line As Variant
    Dim code_template As String
    
    For Each line In Split(m_step_function_template, vbLf)
        If code_template <> vbNullString Then code_template = code_template & vbLf
        code_template = code_template & Trim(line)
    Next
    TSpec.expect(code_template).to_be step_expressions(1)
End Sub

Public Sub Then_the_name_of_the_resulting_function_is_STR_C88DE530F4BA(step_expressions As Collection)
    'Then the name of the resulting function is {string}
    TSpec.expect(m_step_function_name).to_be step_expressions(1)
End Sub

