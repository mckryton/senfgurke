VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TSpecExpectation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mvarGivenValue As Variant

Public Sub to_be(pExpectedValue)

    Dim matcher_fail_msg As String
    
    test_comparability Me.given_value, pExpectedValue
    
    If Not values_are_equal(Me.given_value, pExpectedValue) Then
        If IsObject(Me.given_value) Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                                vbTab & "can't match objects"
            raise_expectation_err matcher_fail_msg
        Else
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                                vbTab & "found:" & vbTab & vbTab & ">" & Me.given_value & "<" & vbLf & _
                                vbTab & "expected:" & vbTab & ">" & pExpectedValue & "<"
            raise_expectation_err matcher_fail_msg
        End If
    End If
End Sub

Public Sub not_to_be(pExpectedValue)

    Dim matcher_fail_msg As String
    
    test_comparability Me.given_value, pExpectedValue
    
    If values_are_equal(Me.given_value, pExpectedValue) Then
        If IsObject(Me.given_value) Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                                vbTab & "compared objects match"
            raise_expectation_err matcher_fail_msg
        Else
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                                vbTab & "found:" & vbTab & vbTab & ">" & Me.given_value & "<" & vbLf & _
                                vbTab & "expected:" & vbTab & "not >" & pExpectedValue & "<"
            raise_expectation_err matcher_fail_msg
        End If
    End If
End Sub

Public Sub to_be_nothing()

    Dim matcher_fail_msg As String
    
    If Not IsObject(Me.given_value) Then
        matcher_fail_msg = vbTab & "expectation failed" & vbLf & _
                                vbTab & "given value:" & vbTab & vbTab & ">" & Me.given_value & "<" & vbLf & _
                                vbTab & "is not an object"
            raise_expectation_err matcher_fail_msg
    Else
        If Not Me.given_value Is Nothing Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                                vbTab & "found an object" & vbLf & _
                                vbTab & "expected:" & vbTab & ">Nothing<"
            raise_expectation_err matcher_fail_msg
        End If
    End If
End Sub

Public Sub not_to_be_nothing()

    Dim matcher_fail_msg As String
    
    If Not IsObject(Me.given_value) Then
        matcher_fail_msg = vbTab & "expectation failed" & vbLf & _
                                vbTab & "given value:" & vbTab & vbTab & ">" & Me.given_value & "<" & vbLf & _
                                vbTab & "is not an object"
            raise_expectation_err matcher_fail_msg
    Else
        If Me.given_value Is Nothing Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                                vbTab & "found nothing" & vbLf & _
                                vbTab & "expected:" & vbTab & ">an object<"
            raise_expectation_err matcher_fail_msg
        End If
    End If
End Sub

Private Sub test_comparability(pGivenValue As Variant, pExpectedValue As Variant)
    
    Dim matcher_fail_msg As String
    
    If (IsObject(Me.given_value) And Not IsObject(pExpectedValue)) _
     Or (IsObject(pExpectedValue) And Not IsObject(Me.given_value)) Then
        matcher_fail_msg = vbTab & "can't compare given with expected value - check if only one of them is an object"
        raise_expectation_err matcher_fail_msg
    End If
End Sub

Private Function values_are_equal(pGivenValue As Variant, pExpectedValue As Variant) As Boolean

    Dim index As Integer
    
    values_are_equal = True
    If IsObject(pGivenValue) Then
        If Not (pGivenValue Is pExpectedValue) Then
            values_are_equal = False
        End If
    ElseIf IsArray(pGivenValue) Then
        If UBound(pGivenValue) <> UBound(pExpectedValue) Then
            values_are_equal = False
        Else
            For index = 0 To UBound(pGivenValue)
                If pGivenValue(index) <> pExpectedValue(index) Then
                   values_are_equal = False
                   Exit For
                End If
            Next
        End If
    Else
        If pGivenValue <> pExpectedValue Then
            values_are_equal = False
        End If
    End If
End Function

Public Property Get given_value() As Variant
    On Error Resume Next
    Set given_value = mvarGivenValue
    given_value = mvarGivenValue
    On Error GoTo 0
End Property

Public Property Let given_value(ByVal pvarGivenValue As Variant)
    On Error Resume Next
    Set mvarGivenValue = pvarGivenValue
    mvarGivenValue = pvarGivenValue
    On Error GoTo 0
End Property

Public Sub contains_member(expected_member As Variant)

    Dim matcher_fail_msg As String
    Dim expected_member_text As String

    If TypeName(Me.given_value) = "Collection" Then
        If Not ExtraVBA.collection_has_value(expected_member, Me.given_value) Then
            If IsArray(expected_member) Then
                expected_member_text = Join(expected_member, "*")
            Else
                expected_member_text = CStr(expected_member)
            End If
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                        vbTab & "collection does not contain item " & ">" & expected_member_text & "<"
            raise_expectation_err matcher_fail_msg
        
        End If
    ElseIf IsArray(Me.given_value) Then
        Debug.Print "TSpecExpectation PENDING: contains_member matcher for array"
    End If
End Sub

Public Sub includes_text(include_item As Variant)

    Dim matcher_fail_msg As String
    
    If TypeName(include_item) = "String" Then
        If InStr(Me.given_value, include_item) = 0 Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                        vbTab & "string:" & vbTab & vbTab & ">" & Me.given_value & "<" & vbLf & _
                        vbTab & "should include:" & vbTab & ">" & include_item & "<"
            raise_expectation_err matcher_fail_msg
        End If
    Else
        Debug.Print "PENDING: include test for arrays"
    End If
End Sub

Public Sub does_not_include(include_item As Variant)

    Dim matcher_fail_msg As String
    
    If TypeName(include_item) = "String" Then
        If InStr(Me.given_value, include_item) > 0 Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                        vbTab & "string:" & vbTab & vbTab & ">" & Me.given_value & "<" & vbLf & _
                        vbTab & "should not include:" & vbTab & ">" & include_item & "<"
            raise_expectation_err matcher_fail_msg
        End If
    Else
        Debug.Print "STATUS_PENDING: include test for arrays"
    End If
End Sub

Private Sub raise_expectation_err(fail_msg)
    
    TSpec.LastFailMsg = fail_msg
    Err.Raise Number:=ERR_ID_EXPECTATION_STATUS_FAILED, Description:=fail_msg
End Sub


