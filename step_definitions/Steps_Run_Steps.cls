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

Public Sub Given_a_list_of_numbers_FF6B22C2A68F(example_context As TContext, data_table As TDataTable)
    ' this is not an actual step but a step inside the example "calling step function for a step with a data_table"
    ' the step is about validating that data_table is set as step function parameter
    ' example_context refers to the context of that example
    example_context.set_value data_table, "data_table"
End Sub

Public Sub Given_a_step_returning_pending_without_a_pending_m_F201A405CFB1(example_context As TContext)
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
    Dim parent_feature As TFeature
    Dim step As TStep
    
    Set parent_feature = THelper.new_TFeature
    parent_feature.parsed_lines = -1
    Set step = TStepParser.parse_step_list(step_expressions(1), parent_feature)(1)
    example_context.set_value step, "example_step"
End Sub

Public Sub When_the_step_is_executed_CEB06B444334(example_context As TContext)
    'When the step is executed
    Dim step_runner As TStepRunner
    Dim session As TSession
    Dim step_classes As Collection
    
    Set step_runner = THelper.new_TStepRunner
    Set step_classes = New Collection
    step_classes.Add New Steps_predefined_steps
    Set session = THelper.new_TSession
    With session
        .reporter.enable_silent_mode
        .StepImplementations = step_classes
    End With
    example_context.set_value session, "session"
    step_runner.setup session
    example_context.set_value step_runner.run_step(example_context.get_value("example_step"), THelper.new_TContext), "step_result"
End Sub

Public Sub When_the_code_template_for_the_step_implementation_7D39D6705DD8(example_context As TContext)
    'When the code template for the step implementation is requested
    example_context.set_value example_context.get_value("example_step").get_step_function_template, "step_function_template"
End Sub

Public Sub When_the_step_is_translated_into_a_function_name_D820185C4AF6(example_context As TContext)
    'When the step is translated into a function name
    example_context.set_value example_context.get_value("example_step").get_step_function_name, "step_function_name"
End Sub

Public Sub When_the_step_function_is_called_418C5F83C560(example_context As TContext)
    'When the step function is called
    Dim step As TStep
    Dim step_runner As TStepRunner
    Dim step_result As Variant
    
    Set step_runner = THelper.new_TStepRunner
    Set step = example_context.get_value("example_step")
    step_runner.try_step Me, step.get_step_function_name, step.Expressions, example_context, step.data_table
    Set step_runner = Nothing
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

Public Sub Then_the_data_table_is_passed_as_a_parameter_CA3BC28813EC(example_context As TContext)
    'Then the data table is passed as a parameter
    Dim data_table As TDataTable
    
    Set data_table = example_context.get_value("data_table")
    TSpec.expect(data_table.column_names.Count).to_be 1
    TSpec.expect(data_table.table_rows.Count).to_be 3
End Sub
