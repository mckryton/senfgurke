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
Dim m_parsed_feature As Collection
Dim m_test_step_implementation As Variant
Dim m_step_result As Variant
Dim m_example_statistics As Collection

Public Property Get Description() As String

    Description = "Senfgurke will locate the matching step implementation for each step" & vbLf _
                   & "of an example and execute it."
End Property

' in the future this property shall be replaced with a function that parses text
Public Property Get Examples() As Collection

    Dim examples_list As Collection
    Dim example As Variant

    Set examples_list = New Collection
    
    example = Array("Rule: step implementations without errors return ""OK"" as status")
    examples_list.Add example
    example = Array("Example: Empty test step", _
                        "Given an example with a step ""Given a step with an empty implementation""", _
                        "And the step has a matching step implementation", _
                        "When the step is executed", _
                        "Then the execution result is ""OK""")
    examples_list.Add example
'    example = Array("Example: test step does match expectation", _
'                        "Given an example with a step ""Given a step with a valid expectation""", _
'                        "And the step has a matching step implementation", _
'                        "When the step is executed", _
'                        "Then the execution result is ""OK""")
'    examples_list.Add example
    
'    example = Array("Rule: step implementations raising an errors return ""FAIL"" as status")
'    examples_list.Add example
'    example = Array("Example: test step does not match expectation", _
'                        "Given an example with a step ""Given a step with an invalid expectation""", _
'                        "And the step has a matching step implementation", _
'                        "When the step is executed", _
'                        "Then the execution result is ""FAIL""")
'    examples_list.Add example

    example = Array("Rule: examples are executed until a step returns a status other than ""OK""")
    examples_list.Add example
    example = Array("Example: example has 3 steps and 2nd fails", _
                        "Given an feature with one example with 3 steps which return OK except the 2nd step fails", _
                        "And all steps have a matching step implementation", _
                        "When the example is executed", _
                        "Then only two steps of the example were executed")
    examples_list.Add example

    Set Examples = examples_list
End Property

Public Function run_step(step_definition As Collection) As Variant

    Dim step_attributes As Collection
    Dim feature_clause As Collection
    
    On Error GoTo error_handler
    Select Case step_definition.Item(TExampleRunner_Old.ATTR_STEP_HEAD) & " " & step_definition.Item(TExampleRunner_Old.ATTR_STEP_BODY)
        
        Case "Given an example with a step ""Given a step with an empty implementation"""
            Set m_parsed_feature = TFeatureParser.parse_feature( _
                "Feature: sample feature" & vbLf & _
                vbLf & _
                "  Example: sample example" & vbLf & _
                "    Given a step with an empty implementation")
                                            
        Case "Given the step has a matching step implementation"
            'because VBA does not support closures this test is using a pre-built step implementation
            Set m_test_step_implementation = New Support_Test_Feature_Class
            
        Case "Given all steps have a matching step implementation"
            'because VBA does not support closures this test is using a pre-built step implementation
            Set m_test_step_implementation = New Support_Test_Feature_Class
            
        Case "Given an feature with one example with 3 steps which return OK except the 2nd step fails"
            Set m_parsed_feature = TFeatureParser.parse_feature( _
                "Feature: sample feature" & vbLf & _
                vbLf & _
                "  Example: sample example" & vbLf & _
                "    Given a step with a valid implementation" & vbLf & _
                "    And a step with an invalid implementation" & vbLf & _
                "    And a step with a valid implementation")
            
        Case "When the step is executed"
            m_step_result = TExampleRunner.execute_step(m_parsed_feature(2)(EXAMPLE_ATTR_STEPS)(1), m_test_step_implementation)
            
        Case "When the example is executed"
            Set m_example_statistics = TExampleRunner.execute_example(m_parsed_feature(2), m_test_step_implementation)
            
        Case "Then the execution result is ""OK"""
            TSpec.expect(m_step_result(0)).to_be "OK"
        
        Case Else
            Err.Raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array(TExampleRunner.STEP_RESULT_OK)
    Exit Function

error_handler:
    run_step = TExampleRunner.fail_step(Err.Number, Err.Description)
End Function

Public Sub before_example()
    
End Sub

Public Sub after_example()
    Set m_test_step_implementation = Nothing
End Sub

Public Property Get tags() As Variant
    tags = cTags
End Property

Private Sub pending(pending_msg)
    TFeatureRunner.pending pending_msg
End Sub

