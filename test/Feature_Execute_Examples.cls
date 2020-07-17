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
Const cTags = ""

Dim m_example As Variant
Dim m_step_attributes As Collection
Dim m_test_feature_object As Variant
Dim m_step_result As Variant

Public Property Get description() As String

    description = "Senfgurke will validate the Gherkin syntax of the given examples" & vbLf _
                   & "and calculate a result for every step of the example."
End Property

Public Sub test_examples()

    Dim examples As Collection
    Dim example As Variant

    Set examples = New Collection
    example = Array("Rule: Steps not starting with Given, When, Then or And will produce a Gherkin syntax error")
    examples.Add example
    example = Array("Example: missing step type", _
                        "Given an example with a step ""bla bla bla""", _
                        "When the step is read", _
                        "Then the step result is an Gherkin Syntax Error claiming ""bla"" is not a valid step type")
    examples.Add example
    
    example = Array("Rule: ""And"" and ""But"" are used as a synonym for the previous step type Given, When or Then")
    examples.Add example
    example = Array("Example: And as synonym for Given", _
                        "Given an example with two steps ""Given x is 1"" and ""And y is 2""", _
                        "When the second step is read", _
                        "Then the step type is set to ""Given""")
    examples.Add example
    example = Array("Example: But as synonym for When", _
                        "Given an example with two steps ""When some action happens"" and ""But it doesn't matter""", _
                        "When the second step is read", _
                        "Then the step type is set to ""When""")
    examples.Add example
    
    example = Array("Rule: implemented test steps are OK if no error was raised, otherwise the steps FAIL")
    examples.Add example
    example = Array("Example: Empty test step", _
                        "Given an example with a step ""Given a step with an empty implementation""", _
                        "And the step has a matching step implementation", _
                        "When the step is executed", _
                        "Then the execution result is ""OK""")
    examples.Add example
    example = Array("Example: test step does match expectation", _
                        "Given an example with a step ""Given a step with a valid expectation""", _
                        "And the step has a matching step implementation", _
                        "When the step is executed", _
                        "Then the execution result is ""OK""")
    examples.Add example
    example = Array("Example: test step does not match expectation", _
                        "Given an example with a step ""Given a step with an invalid expectation""", _
                        "And the step has a matching step implementation", _
                        "When the step is executed", _
                        "Then the execution result is ""FAIL""")
    examples.Add example
    
    
'    example = Array("Rule: Every step will result in one of the states OK, FAILED, MISSING or PENDING")
'    examples.Add example
'
'    example = Array("Rule: FAILED, MISSING and PENDING test steps will end the example execution and continue with the next example")
'    examples.Add example
    
    For Each example In examples
        TExampleRunner.run_example example, Me
    Next
End Sub

Public Function run_step(pcolStepProps As Collection) As Variant

    Dim step_attributes As Collection
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item(TExampleRunner.ATTR_STEP_HEAD) & " " & pcolStepProps.Item(TExampleRunner.ATTR_STEP_BODY)
    
        Case "Given an example with a step ""bla bla bla"""
            m_example = Array("Example: Gherkin syntax error", _
                                "bla bla bla")
                        
        Case "Given an example with two steps ""Given x is 1"" and ""And y is 2"""
            m_example = Array("Example: Using And as Given", _
                                "Given x is 1", _
                                "And y is 2")
        
        Case "Given an example with two steps ""When some action happens"" and ""But it doesn't matter"""
            m_example = Array("Example: Using But as When", _
                                "When some action happens", _
                                "But it doesn't matter")
        
        Case "Given an example with a step ""Given a step with an empty implementation"""
             m_example = Array("Example: Empty test step", _
                                "Given a step with an empty implementation")
            
        Case "Given an example with a step ""Given a step with a valid expectation"""
            m_example = Array("Example: Matching expectation", _
                                "Given a step with a valid expectation")
        
        Case "Given an example with a step ""Given a step with an invalid expectation"""
            m_example = Array("Example: Failing expectation", _
                                "Given a step with an invalid expectation")
                                
        Case "Given the step has a matching step implementation"
            Set m_test_feature_object = New Support_Execute_Steps
        
        Case "When the second step is read"
            Set m_step_attributes = TExampleRunner.read_step(m_example, 2)
        
        Case "When the step is read"
            Set m_step_attributes = TExampleRunner.read_step(m_example, 1)
            
        Case "When the step is executed"
            Set step_attributes = TExampleRunner.read_step(m_example, 1)
            m_step_result = TExampleRunner.execute_step(step_attributes, m_test_feature_object)
        
        Case "Then the step result is an Gherkin Syntax Error claiming ""bla"" is not a valid step type"
            TSpec.expect(m_step_attributes(TExampleRunner.ATTR_ERROR)).to_be "Gherkin syntax error: >bla< is not a valid step type"
        
        Case "Then the step type is set to ""Given"""
            TSpec.expect(m_step_attributes(TExampleRunner.ATTR_STEP_HEAD)).to_be "Given"
            
        Case "Then the step type is set to ""When"""
            TSpec.expect(m_step_attributes(TExampleRunner.ATTR_STEP_HEAD)).to_be "When"
            
        Case "Then the execution result is ""OK"""
            TSpec.expect(m_step_result(0)).to_be "OK"
        
        Case Else
            Err.Raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array("OK")
    Exit Function

error_handler:
    run_step = TFeature.fail_step(Err.Number, Err.description)
End Function

Public Sub before_example()
    
End Sub

Public Sub after_example()
    Set m_test_feature_object = Nothing
    Set m_step_attributes = Nothing
End Sub

Public Property Get tags() As Variant
    tags = cTags
End Property

Private Sub pending(pending_msg)
    TFeature.pending pending_msg
End Sub

Private Property Get log() As Logger
    Set log = TFeature.log
End Property



