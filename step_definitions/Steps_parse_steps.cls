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

Public Sub Given_an_example_with_a_step_STR_2AA83F29F0BD(example_context As TContext, step_expressions As Collection)
    'Given an example with a step {string}
    example_context.set_value New TExample, "example"
    example_context.get_value("example").Steps.Add TFeatureParser.create_step(step_expressions(1))
End Sub

Public Sub Given_another_step_is_defined_as_STR_F14E826F5231(example_context As TContext, step_expressions As Collection)
    'Given another step is defined as {string}
    example_context.set_value step_expressions(1), "gherkin_text"
End Sub

Public Sub When_the_step_definition_is_parsed_8240391A8855(example_context As TContext)
    'When the step definition is parsed
    example_context.set_value TFeatureParser.create_step(example_context.get_value("gherkin_text")), "step"
End Sub

Public Sub When_the_step_definition_is_parsed_as_part_of_this_E03DC3C99405(example_context As TContext)
    'When the step definition is parsed as part of this example
    example_context.set_value TFeatureParser.create_step(example_context.get_value("gherkin_text"), example_context.get_value("example")), "step"
End Sub

Public Sub Then_the_type_of_the_step_is_set_to_STR_3EC12192CDA4(example_context As TContext, step_expressions As Collection)
    'Then the type of the step is set to {string}
    TSpec.expect(example_context.get_value("step").SType).to_be step_expressions(1)
End Sub

Public Sub Then_the_name_of_the_step_is_STR_47BC4E071438(example_context As TContext, step_expressions As Collection)
    'Then the name of the step is "a step"
    TSpec.expect(example_context.get_value("step").name).to_be step_expressions(1)
End Sub

