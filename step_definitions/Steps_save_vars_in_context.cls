VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_save_vars_in_context"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub When__STR_is_added_as_STR_in_context_6DEAADD7C784(example_context As TContext, step_expressions As Collection)
    'When  {string} is added as {string} in context
    example_context.set_value step_expressions(1), step_expressions(2)
End Sub

Public Sub When__INT_is_add_as_STR_in_context_715E61B1C784(example_context As TContext, step_expressions As Collection)
    'When  {integer} is add as {string} in context
    example_context.set_value step_expressions(1), step_expressions(2)
End Sub

Public Sub When_a_new_collection_is_added_as_STR_to_the__42F4D1BED80C(example_context As TContext, step_expressions As Collection)
    'When a new collection is added as {string} to the context
    example_context.set_value New Collection, step_expressions(1)
End Sub

Public Sub Then_the_context_value_STR_returns_STR_2D0DA156E7CD(example_context As TContext, step_expressions As Collection)
    'Then the context value {string} returns {string}
    TSpec.expect(example_context.get_value(step_expressions(1))).to_be step_expressions(2)
End Sub

Public Sub Then_the_context_value_STR_returns_INT_2D0DA1564267(example_context As TContext, step_expressions As Collection)
    'Then the context value {string} returns {integer}
    TSpec.expect(example_context.get_value(step_expressions(1))).to_be step_expressions(2)
End Sub

Public Sub Then_the_context_value_STR_returns_a_collecti_0B0B7BAA0497(example_context As TContext, step_expressions As Collection)
    'Then the context value {string} returns a collection object
    TSpec.expect(TypeName(example_context.get_value(step_expressions(1)))).to_be "Collection"
End Sub
