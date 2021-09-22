VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_parse_steps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_a_step_definition_STR_65172AC6DD8A(example_context As TContext, step_expressions As Collection)
    'Given a step definition {string}
    example_context.set_value step_expressions(1), "gherkin_text"
End Sub

Public Sub Given_a_step_is_defined_as_STR_2925BEC1427D(example_context As TContext, step_expressions As Collection)
    'Given a step is defined as "Given a step"
    example_context.set_value step_expressions(1), "gherkin_text"
End Sub

Public Sub Given_a_list_of_steps_STR_D7FEE94DD01B(example_context As TContext, step_expressions As Collection)
    'Given a list of steps {string}
    'TODO refactor stepexpression to remove escape characters the right way
    Dim step_list_gherkin As String
    
    step_list_gherkin = Replace(step_expressions(1), "\", "")
    example_context.set_value step_list_gherkin, "gherkin_text"
End Sub

Public Sub When_the_step_definition_is_parsed_8240391A8855(example_context As TContext)
    'When the step definition is parsed
    example_context.set_value TStepParser.parse_step_line(example_context.get_value("gherkin_text"), New Collection), "step"
End Sub

Public Sub When_the_step_list_is_parsed_8C8F753FF01F(example_context As TContext)
    'When the step list is parsed
    Dim parent_feature As TFeature

    Set parent_feature = New TFeature
    'parser should have parsed the feature text up to the line before the step list
    parent_feature.parsed_lines = -1
    example_context.set_value TStepParser.parse_step_list(example_context.get_value("gherkin_text"), parent_feature), "step_list"
End Sub

Public Sub When_the_type_of_the_step_line_is_evaluated_ABBCD0B4A184(example_context As TContext)
    'When the type of the step line is evaluated
    example_context.set_value TFeatureParser.get_line_type(example_context.get_value("gherkin_text")), "line_type"
End Sub

Public Sub Then_the_type_of_the_step_is_set_to_STR_3EC12192CDA4(example_context As TContext, step_expressions As Collection)
    'Then the type of the step is set to {string}
    TSpec.expect(example_context.get_value("step").SType).to_be step_expressions(1)
End Sub

Public Sub Then_the_name_of_the_step_is_STR_47BC4E071438(example_context As TContext, step_expressions As Collection)
    'Then the name of the step is "a step"
    TSpec.expect(example_context.get_value("step").name).to_be step_expressions(1)
End Sub

Public Sub Then_the_type_of_step_INT_is_set_to_STR_3EC1B545CDA4(example_context As TContext, step_expressions As Collection)
    'Then the type of step {integer} is set to {string}
    Dim step_list As Collection
    Dim step As TStep
    
    Set step_list = example_context.get_value("step_list")
    Set step = step_list(step_expressions(1) - 1)
    TSpec.expect(step.SType).to_be step_expressions(2)
End Sub

Public Sub Then_the_resulting_line_type_is_not_STR_4C0165DE378D(example_context As TContext, step_expressions As Collection)
    'Then the resulting line type is not {string}
    TSpec.expect(example_context.get_value("line_type")).not_to_be step_expressions(1)
End Sub

Public Sub Then_the_resulting_step_list_contains_only_INT_ste_DD326A098DF8(example_context As TContext, step_expressions As Collection)
    'Then the resulting step list contains only {integer} steps
    TSpec.expect(example_context.get_value("step_list").Count).to_be step_expressions(1)
End Sub

Public Sub Then_the_name_of_step_INT_is_STR_47BC4FBE3041(example_context As TContext, step_expressions As Collection)
    'And the name of step {integer} is {string}
    Dim step_list As Collection
    Dim step As TStep
    
    Set step_list = example_context.get_value("step_list")
    Set step = step_list(step_expressions(1))
    TSpec.expect(step.name).to_be step_expressions(2)
End Sub
