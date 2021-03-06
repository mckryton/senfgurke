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

Private m_collection As Collection
Private m_expected_value As Variant
Private m_expectation_result As Boolean

Public Sub Given_a_collection_has_INT_members_array_INT__STR__22433998BB74(step_expressions As Collection)
    'Given a collection has {integer} members array( {integer}, {string} ) and array( {integer}, {string} )
    Dim index As Integer
    
    Set m_collection = New Collection
    For index = 1 To step_expressions(1) * 2 Step 2
        m_collection.Add Array(step_expressions(index + 1), step_expressions(index + 2))
    Next
End Sub

Public Sub Given_an_expected_value_is_an_array_INT__STR__C2EB84347C9F(step_expressions As Collection)
    'And an expected value is an array( {integer}  {string} )
    m_expected_value = Array(step_expressions(1), step_expressions(2))
End Sub

Public Sub When_the_expectation_validates_that_the_collection_808FE2AA1406()
    'When the expectation validates that the collection contains the expected value
    m_expectation_result = True
    On Error GoTo expectation_failed
    TSpec.expect(m_collection).contains_member m_expected_value
    Exit Sub
    
expectation_failed:
    m_expectation_result = False
End Sub

Public Sub Then_the_expectation_is_confirmed_A341C2A69D10()
    'Then the expectation is confirmed
    TSpec.expect(m_expectation_result).to_be True
End Sub
