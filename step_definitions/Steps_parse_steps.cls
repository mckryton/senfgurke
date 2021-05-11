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

Public Sub Given_a_step_definition_STR_65172AC6DD8A(context As TContext, step_expressions As Collection)
    'Given a step definition {string}
    TStepVars.gherkin_text = step_expressions(1)
End Sub

Public Sub Given_a_step_is_defined_as_STR_2925BEC1427D(context As TContext, step_expressions As Collection)
    'Given a step is defined as "Given a step"
    TStepVars.gherkin_text = step_expressions(1)
End Sub

Public Sub Given_an_example_with_a_step_STR_2AA83F29F0BD(context As TContext, step_expressions As Collection)
    'Given an example with a step {string}
    Set TStepVars.example = New TExample
    TStepVars.example.Steps.Add TFeatureParser.create_step(step_expressions(1))
End Sub

Public Sub Given_another_step_is_defined_as_STR_F14E826F5231(context As TContext, step_expressions As Collection)
    'Given another step is defined as {string}
    TStepVars.gherkin_text = step_expressions(1)
End Sub

Public Sub When_the_step_definition_is_parsed_8240391A8855(context As TContext)
    'When the step definition is parsed
    Set TStepVars.step = TFeatureParser.create_step(TStepVars.gherkin_text)
End Sub

Public Sub When_the_step_definition_is_parsed_as_part_of_this_E03DC3C99405(context As TContext)
    'When the step definition is parsed as part of this example
    Set TStepVars.step = TFeatureParser.create_step(TStepVars.gherkin_text, TStepVars.example)
End Sub

Public Sub Then_the_type_of_the_step_is_set_to_STR_3EC12192CDA4(context As TContext, step_expressions As Collection)
    'Then the type of the step is set to {string}
    TSpec.expect(TStepVars.step.SType).to_be step_expressions(1)
End Sub

Public Sub Then_the_name_of_the_step_is_STR_47BC4E071438(context As TContext, step_expressions As Collection)
    'Then the name of the step is "a step"
    TSpec.expect(TStepVars.step.name).to_be step_expressions(1)
End Sub

