Attribute VB_Name = "TSpec"
Option Explicit

Dim m_last_fail_msg As String

Public Function expect(pvarGivenValue As Variant) As TSpecExpectation

    Dim expectation As TSpecExpectation
    
    Set expectation = New TSpecExpectation
    expectation.given_value = pvarGivenValue
    Set expect = expectation
End Function

Public Property Get LastFailMsg() As String
    
    LastFailMsg = m_last_fail_msg
End Property

Public Property Let LastFailMsg(ByVal last_fail_msg As String)
    ' this is a workaround for VBAs mishandling of custom errors
    '  e.g. using callbyname() will overwritecustom errors with 440 automation error
    '   https://stackoverflow.com/questions/18241906/vba-callbyname-and-invokehook
    '  e.g. using strongly typed objects will overwrite the error description
    '   https://stackoverflow.com/questions/31234805/err-raise-is-ignoring-custom-description-and-source
    ' so in addition of adding expectation results to custom errors (aka exceptions), the most current result meessage will be saved here

    m_last_fail_msg = last_fail_msg
End Property
