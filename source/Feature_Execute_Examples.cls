VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feature_Execute_Examples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'separate tags by comma, e.g.: "wip,debug,production"
Const cTags = "wip"

Dim m_example As Variant
Dim m_example_attributes As Collection

Public Property Get description() As String

    description = "Senfgurke will validate the Gherkin syntax of the given examples" & vbLf & _
                    vbTab & "and calculate a result for every step of the example."
End Property

Public Sub test_examples()

    Dim examples As Collection
    Dim example As Variant

    Set examples = New Collection
    example = Array("Rule: Steps not starting with Given, When, Then or And will produces a Gherkin syntax error")
    examples.Add example
    example = Array("Example: missing step type", _
                        "Given an example with a step ""bla bla bla""", _
                        "When the step is read", _
                        "Then the step result is an Gherkin Syntax Error claiming ""bla"" is not a valid step type")
    examples.Add example
    example = Array("Rule: ""And"" is used as a synonym for the previous step type Given, When or Then")
    examples.Add example
    example = Array("Rule: Examples using Given, When, Then in a different order will produce a Gherkin Syntax error")
    examples.Add example
    example = Array("Rule: Every step will result in one of the states OK, FAILED, MISSING or PENDING")
    examples.Add example
    example = Array("Rule: FAILED, MISSING and PENDING test steps will end the example excution and continue with the next example")
    examples.Add example
    
    For Each example In examples
        TExampleRunner.run_example example, Me
    Next
End Sub

Public Function run_step(pcolStepProps As Collection) As Variant
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item(TExampleRunner.ATTR_STEP_HEAD) & " " & pcolStepProps.Item(TExampleRunner.ATTR_STEP_BODY)
    
        Case "Given an example with a step ""bla bla bla"""
            m_example = Array("Example: Gherkin syntax error", _
                        "bla bla bla")
        
        Case "When the step is read"
            Set m_example_attributes = TExampleRunner.read_example_step(m_example, 1)
        
        Case "Then the step result is an Gherkin Syntax Error claiming ""bla"" is not a valid step type"
            TSpec.expect(m_example_attributes(TExampleRunner.ATTR_ERROR)).to_be "Gherkin syntax error: >bla< is not a valid step type"
        
        Case Else
            Err.raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array("OK")
    Exit Function

error_handler:
    run_step = TFeature.raise_step_error(Err.Number, Err.description)
End Function

Public Sub before()
    
End Sub

Public Sub after()

End Sub

Public Property Get Tags() As Variant
    Tags = cTags
End Property

Private Sub pending(pending_msg)
    TFeature.pending pending_msg
End Sub

Private Property Get Log() As Logger
    Set Log = TFeature.Log
End Property



