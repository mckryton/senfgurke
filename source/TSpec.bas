Attribute VB_Name = "TSpec"
Option Explicit

Public Function expect(pvarGivenValue As Variant) As TSpecExpectation

    Dim expectation As TSpecExpectation
    
    Set expectation = New TSpecExpectation
    expectation.given_value = pvarGivenValue
    Set expect = expectation
End Function

