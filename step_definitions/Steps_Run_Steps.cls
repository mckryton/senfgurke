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

Public Sub Given_a_pending_step_6A3533BBBFE5(example_context As TContext)
    pending
End Sub

Public Sub Given_an_empty_step_6A35DCCE89C0(example_context As TContext)
    'Given an empty step
    
End Sub

Public Sub Given_a_valid_step_8A74152FD2F9(example_context As TContext)
    TSpec.expect(1).to_be 1
End Sub

Public Sub Given_an_invalid_step_CF4B201BCF1B(example_context As TContext)
    TSpec.expect(1).to_be 2
End Sub

Public Sub Given_a_step_returning_PENDING_without_a_pending_m_F201CAEBCFB1(example_context As TContext)
    'Given a step returning PENDING without a pending message
    example_context.set_value TStepParser.parse_step_line("Given a pending step", New Collection), "example_step"
    'won't create code dynamically, so step "Given a pending step" is hard coded
End Sub

Public Sub Given_a_step_with_an_empty_implementation_BC7B46584C5E(example_context As TContext)
    'Given a step with an empty implementation
    example_context.set_value TStepParser.parse_step_line("Given an empty step", New Collection), "example_step"
    'the step function itself can't be created dynamically and has to be available
End Sub

Public Sub Given_a_step_with_a_working_implementation_05BC882CCAAA(example_context As TContext)
    'Given a step with a working implementation
    example_context.set_value TStepParser.parse_step_line("Given a valid step", New Collection), "example_step"
    'the step function itself can't be created dynamically and has to be available
End Sub

Public Sub Given_a_step_with_a_failing_implementation_05BCA3B5CAAA(example_context As TContext)
    'Given a step with a failing implementation
    example_context.set_value TStepParser.parse_step_line("Given an invalid step", New Collection), "example_step"
    'the step function itself can't be created dynamically and has to be available
End Sub

Public Sub Given_a_step_without_an_implementation_FAF18488CAAA(example_context As TContext)
    'Given a step without an implementation
    example_context.set_value TStepParser.parse_step_line("Given a random step name " & CStr(Now), New Collection), "example_step"
End Sub

Public Sub Given_a_step_STR_F5C6F4BA1D15(example_context As TContext, step_expressions As Collection)
    'Given a step {string}
    example_context.set_value TStepParser.parse_step_line(step_expressions(1), New Collection), "example_step"
End Sub

Public Sub When_the_step_is_executed_CEB06B444334(example_context As TContext)
    'When the step is executed
    Dim step_runner As TStepRunner
    
    Set step_runner = New TStepRunner
    example_context.set_value New TSession, "session"
    example_context.get_value("session").Reporter.enable_silent_mode
    step_runner.setup example_context.get_value("session")
    example_context.set_value step_runner.run_step(example_context.get_value("example_step"), New TContext), "step_result"
End Sub

Public Sub When_the_code_template_for_the_step_implementation_7D39D6705DD8(example_context As TContext)
    'When the code template for the step implementation is requested
    example_context.set_value example_context.get_value("example_step").get_step_function_template, "step_function_template"
End Sub

Public Sub When_the_step_is_translated_into_a_function_name_D820185C4AF6(example_context As TContext)
    'When the step is translated into a function name
    example_context.set_value example_context.get_value("example_step").get_step_function_name, "step_function_name"
End Sub

Public Sub Then_the_execution_result_is_STR_80FDCF952A5A(example_context As TContext, step_expressions As Collection)
    'Then the execution result is {string}
    TSpec.expect(example_context.get_value("step_result")(0)).to_be step_expressions(1)
End Sub

Public Sub Then_the_code_template_for_the_step_implementation_9C79C155DB8E(example_context As TContext, step_expressions As Collection)
    'Then the code template for the step implementation is {string}
    TSpec.expect(example_context.get_value("step_function_template")).to_be step_expressions(1)
End Sub

Public Sub Then_the_name_of_the_resulting_function_is_STR_C88DE530F4BA(example_context As TContext, step_expressions As Collection)
    'Then the name of the resulting function is {string}
    TSpec.expect(example_context.get_value("step_function_name")).to_be step_expressions(1)
End Sub

