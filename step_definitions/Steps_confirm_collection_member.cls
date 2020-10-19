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

Dim m_collection As Collection
Dim m_expected_value As Variant
Dim m_expectation_result As Boolean

Public Sub Given_a_collection_has_2_members_array1a_and_array_275F9A660579()
    'Given a collection has 2 members array(1,"a") and array(2,"b")
    Set m_collection = New Collection
    m_collection.Add Array(1, "a")
    m_collection.Add Array(2, "b")
End Sub

Public Sub Given_an_expected_value_is_an_array2b_575C1E51C6E9()
    'Given an expected value is an array(2,"b")
    m_expected_value = Array(2, "b")
End Sub

Public Sub When_the_expectation_validates_that_the_collection_E065A6F9253F()
    'When the expectation validates that the collection contains the expected value
    m_expectation_result = False
    On Error Resume Next
    m_expectation_result = TSpec.expect(m_collection).contains(m_expected_value)
    On Error GoTo 0
End Sub

Public Sub Then_the_expectation_is_confirmed_BCFEF2868622()
    'Then the expectation is confirmed
    TSpec.expect(m_expectation_result).to_be True
End Sub
