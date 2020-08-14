VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Run_Examples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'separate tags by comma, e.g.: "wip,debug,production"
Const cTags = ""

Dim m_example As Variant
Dim m_test_feature_object As Variant
Dim m_step_result As Variant

Public Property Get Description() As String

    Description = "Senfgurke will validate the Gherkin syntax of the given examples" & vbLf _
                   & "and calculate a result for every step of the example."
End Property

' in the future this property shall be replaced with a function that parses text
Public Property Get Examples() As Collection

    Dim examples_list As Collection
    Dim example As Variant

    Set examples_list = New Collection
    
    example = Array("Rule: implemented test steps are OK if no error was raised, otherwise the steps FAIL")
    examples_list.Add example
    example = Array("Example: Empty test step", _
                        "Given an example with a step ""Given a step with an empty implementation""", _
                        "And the step has a matching step implementation", _
                        "When the step is executed", _
                        "Then the execution result is ""OK""")
    examples_list.Add example
    example = Array("Example: test step does match expectation", _
                        "Given an example with a step ""Given a step with a valid expectation""", _
                        "And the step has a matching step implementation", _
                        "When the step is executed", _
                        "Then the execution result is ""OK""")
    examples_list.Add example
    example = Array("Example: test step does not match expectation", _
                        "Given an example with a step ""Given a step with an invalid expectation""", _
                        "And the step has a matching step implementation", _
                        "When the step is executed", _
                        "Then the execution result is ""FAIL""")
    examples_list.Add example
    
    
'    example = Array("Rule: Every step will result in one of the states OK, FAILED, MISSING or PENDING")
'    examples_list.Add example
'
'    example = Array("Rule: FAILED, MISSING and PENDING test steps will end the example execution and continue with the next example")
'    examples_list.Add example

    Set Examples = examples_list
End Property

Public Function run_step(pcolStepProps As Collection) As Variant

    Dim step_attributes As Collection
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item(TExampleRunner.ATTR_STEP_HEAD) & " " & pcolStepProps.Item(TExampleRunner.ATTR_STEP_BODY)
        
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
            Set m_test_feature_object = New Support_Test_Feature_Class
            
        Case "When the step is executed"
            Set step_attributes = TExampleRunner.read_step(m_example, 1)
            m_step_result = TExampleRunner.execute_step(step_attributes, m_test_feature_object)
            
        Case "Then the execution result is ""OK"""
            TSpec.expect(m_step_result(0)).to_be "OK"
        
        Case Else
            Err.Raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array("OK")
    Exit Function

error_handler:
    run_step = TFeature.fail_step(Err.Number, Err.Description)
End Function

Public Sub before_example()
    
End Sub

Public Sub after_example()
    Set m_test_feature_object = Nothing
End Sub

Public Property Get tags() As Variant
    tags = cTags
End Property

Private Sub pending(pending_msg)
    TFeature.pending pending_msg
End Sub



