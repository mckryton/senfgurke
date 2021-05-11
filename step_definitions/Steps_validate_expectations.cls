VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_validate_expectations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_collection As Collection
Private m_expected_value As Variant
Private m_expectation_result As Boolean
Private m_object_value As Collection
Private m_non_object_value As Variant

Public Sub Given_a_collection_has_INT_members_array_INT__STR__22433998BB74(example_context As TContext, step_expressions As Collection)
    'Given a collection has {integer} members array( {integer}, {string} ) and array( {integer}, {string} )
    Dim index As Integer
    
    Set m_collection = New Collection
    For index = 1 To step_expressions(1) * 2 Step 2
        m_collection.Add Array(step_expressions(index + 1), step_expressions(index + 2))
    Next
End Sub

Public Sub Given_an_expected_value_is_an_array_INT__STR__C2EB84347C9F(example_context As TContext, step_expressions As Collection)
    'And an expected value is an array( {integer}  {string} )
    m_expected_value = Array(step_expressions(1), step_expressions(2))
End Sub

Public Sub Given_a_variable_refers_to_an_object_C82E72756271(example_context As TContext)
    'Given a variable refers to an object
    Set m_object_value = New Collection
End Sub

Public Sub Given_a_variable_refers_to_nothing_6F7619677868(example_context As TContext)
    'Given a variable refers to nothing
    Set m_object_value = Nothing
End Sub

Public Sub Given_a_variable_refers_to_INT_B6E787C91FDE(example_context As TContext, step_expressions As Collection)
    'Given a variable refers to {integer}
    m_non_object_value = step_expressions(1)
End Sub

Public Sub When_the_expectation_validates_that_the_collection_808FE2AA1406(example_context As TContext)
    'When the expectation validates that the collection contains the expected value
    m_expectation_result = True
    On Error GoTo expectation_failed
    TSpec.expect(m_collection).contains_member m_expected_value
    Exit Sub
    
expectation_failed:
    m_expectation_result = False
End Sub

Public Sub When_the_variable_is_expected_to_be_nothing_DBE6C575C088(example_context As TContext)
    'When the variable is expected to be nothing
    m_expectation_result = True
    On Error GoTo expectation_failed
    TSpec.expect(m_object_value).to_be_nothing
    Exit Sub
    
expectation_failed:
    m_expectation_result = False
End Sub

Public Sub When_the_nonobject_variable_is_expected_to_be_noth_10954F3DA4DC(example_context As TContext)
    'When the non-object variable is expected to be nothing
    m_expectation_result = True
    On Error GoTo expectation_failed
    TSpec.expect(m_non_object_value).to_be_nothing
    Exit Sub
    
expectation_failed:
    m_expectation_result = False
End Sub

Public Sub When_the_variable_is_expected_not_to_be_nothing_63DB3DDBA289(example_context As TContext)
    'When the variable is expected not to be nothing
    m_expectation_result = True
    On Error GoTo expectation_failed
    TSpec.expect(m_object_value).not_to_be_nothing
    Exit Sub
    
expectation_failed:
    m_expectation_result = False
End Sub

Public Sub When_the_nonobject_variable_is_expected_not_to_be__A409ACACC890(example_context As TContext)
    'When the non-object variable is expected not to be nothing
    m_expectation_result = True
    On Error GoTo expectation_failed
    TSpec.expect(m_non_object_value).to_be_nothing
    Exit Sub
    
expectation_failed:
    m_expectation_result = False
End Sub


Public Sub Then_the_expectation_is_confirmed_A341C2A69D10(example_context As TContext)
    'Then the expectation is confirmed
    TSpec.expect(m_expectation_result).to_be True
End Sub

Public Sub Then_the_expectation_fails_DBF0F873A79F(example_context As TContext)
    'Then the expectation fails
    TSpec.expect(m_expectation_result).to_be False
End Sub

