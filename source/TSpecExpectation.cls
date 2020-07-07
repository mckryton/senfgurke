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
                                vbTab & "expected:" & vbTab & ">" & pExpectedValue & "<"
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

Public Sub contains(pvarExpectedItem As Variant)

    Dim matcher_fail_msg As String
    
    If TypeName(Me.given_value) = "Collection" Then
        If Not ExtraVBA.existsItem(pvarExpectedItem(0), Me.given_value) Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                        vbTab & "collection does not contain item " & ">" & pvarExpectedItem(0) & "<"
            raise_expectation_err matcher_fail_msg
        ElseIf Me.given_value(pvarExpectedItem(0)) <> pvarExpectedItem(1) Then
            matcher_fail_msg = vbTab & "missed expectation" & vbLf & _
                        vbTab & "found:" & vbTab & vbTab & ">" & Me.given_value(pvarExpectedItem(0)) & "<" & vbLf & _
                        vbTab & "expected:" & vbTab & ">" & pvarExpectedItem(1) & "<"
            raise_expectation_err matcher_fail_msg
        End If
    ElseIf IsArray(Me.given_value) Then
        TExampleRunner.pending "contains matcher for array"
    End If
End Sub

Private Sub raise_expectation_err(pFailMsg)
    
    Err.raise Number:=TErrorDef.ERR_ID_EXPECTATION_FAILED, description:=pFailMsg
End Sub

