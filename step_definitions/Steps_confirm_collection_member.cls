VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_confirm_collection_member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_a_collection_has_INT_members_array_INT__22433998BB74(example_context As TContext, step_expressions As Collection)
    'Given a collection has {integer} members array( {integer}, {string} ) and array( {integer}, {string} )
    Dim index As Integer
    Dim sample_collection As Collection

    Set sample_collection = New Collection
    example_context.set_value sample_collection, "sample_collection"
    For index = 1 To step_expressions(1) * 2 Step 2
        sample_collection.Add Array(step_expressions(index + 1), step_expressions(index + 2))
    Next
End Sub

Public Sub Given_an_expected_value_is_an_array_INT__STR__C2EB84347C9F(example_context As TContext, step_expressions As Collection)
    'And an expected value is an array( {integer}  {string} )
    Dim expected_value As Variant
    
    expected_value = Array(step_expressions(1), step_expressions(2))
    example_context.set_value expected_value, "expected_value"
End Sub

Public Sub When_the_expectation_validates_that_the_colle_808FE2AA1406(example_context As TContext)
    'When the expectation validates that the collection contains the expected value
    Dim sample_collection As Collection
    Dim expected_value As Variant

    Set sample_collection = example_context.get_value("sample_collection")
    expected_value = example_context.get_value("expected_value")
    On Error GoTo expectation_failed
    TSpec.expect(sample_collection).contains_member expected_value
    example_context.set_value "confirmed", "expectation_result"
    Exit Sub
    
expectation_failed:
    example_context.set_value "failed", "expectation_result"
End Sub

