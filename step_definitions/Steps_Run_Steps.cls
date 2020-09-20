VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Run_Steps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'separate tags by comma, e.g.: "wip,debug,production"
Const cTags = "wip"

Dim m_example_step As TStep
Dim m_step_result As Variant
Dim step_implementation_name As String

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
                        "Given a step with an empty implementation", _
                        "When the step is executed", _
                        "Then the execution result is ""OK""")
    examples_list.Add example
    example = Array("Example: test step does match expectation", _
                        "Given a step with a working implementation", _
                        "When the step is executed", _
                        "Then the execution result is ""OK""")
    examples_list.Add example
    
    example = Array("Rule: step implementations raising an errors return ""FAIL"" as status")
    examples_list.Add example
    example = Array("Example: test step does not match expectation", _
                        "Given a step with a failing implementation", _
                        "When the step is executed", _
                        "Then the execution result is ""FAIL""")
    examples_list.Add example
    
    example = Array("Rule: steps without any implementation return ""MISSING"" as status")
    'TODO: add code snippet for the implementation to the result
    examples_list.Add example
    example = Array("Example: test step does not match any implemenation", _
                        "Given a step without an implementation", _
                        "When the step is executed", _
                        "Then the execution result is ""MISSING""")
    examples_list.Add example
    
    
    example = Array("Rule: the name of the step implemenation function is the full step name " & _
                            " stripped from illegal characters and spaces replaced with underscore")
    examples_list.Add example
    example = Array("Example: simple step name", _
                        "Given a step ""Given a step""", _
                        "When the step is translated into a function name", _
                        "Then the name of the resulting function is ""Given_a_step""")
    examples_list.Add example
    'TODO: add more examples after support for scenario outlines is added

    Set Examples = examples_list
End Property

Public Function run_step(step_definition As String) As Variant

    Dim step_attributes As Collection
    Dim feature_clause As Collection
    
    On Error GoTo error_handler
    Select Case step_definition
        
        Case "Given a step with an empty implementation"
            Set m_example_step = New TStep
            With m_example_step
                .SType = STEP_TYPE_GIVEN
                .Name = "an empty step"
            End With
            
        Case "Given an empty step"
            'until VBA adds support for closures, there seems to be no way to create functions on the fly
            ' therefore this step implementation is fix
            
        Case "Given a step with a working implementation"
            Set m_example_step = New TStep
            With m_example_step
                .SType = STEP_TYPE_GIVEN
                .Name = "a valid step"
            End With
            
        Case "Given a valid step"
            TSpec.expect(1).to_be 1
            
        Case "Given a step with a failing implementation"
            Set m_example_step = New TStep
            With m_example_step
                .SType = STEP_TYPE_GIVEN
                .Name = "an invalid step"
            End With
            
        Case "Given an invalid step"
            TSpec.expect(1).to_be 2
            
        Case "Given a step without an implementation"
            Set m_example_step = New TStep
            With m_example_step
                .SType = STEP_TYPE_GIVEN
                .Name = CStr(Now()) 'create a random step name
            End With
                                                                      
        Case "Given a step ""Given a step"""
            Set m_example_step = New TStep
            With m_example_step
                .SType = STEP_TYPE_GIVEN
                .Name = "a step"
            End With
        
        Case "When the step is executed"
            m_step_result = TStepRunner.execute_step(m_example_step)
        
        Case "When the step is translated into a function name"
            step_implementation_name = m_example_step.get_step_function_name()
        
        Case "Then the execution result is ""OK"""
            TSpec.expect(m_step_result(0)).to_be "OK"
            
        Case "Then the execution result is ""FAIL"""
            TSpec.expect(m_step_result(0)).to_be "FAIL"
            
        Case "Then the execution result is ""MISSING"""
            TSpec.expect(m_step_result(0)).to_be "MISSING"
        
        Case "Then the name of the resulting function is ""Given_a_step"""
            TSpec.expect(step_implementation_name).to_be "Given_a_step"
        
        Case Else
            Err.Raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array(STEP_RESULT_OK)
    Exit Function

error_handler:
    run_step = TStepRunner.fail_step(Err.Number, Err.Description)
End Function

Public Sub before_example()
    
End Sub

Public Sub after_example()

End Sub

Public Property Get tags() As Variant
    tags = cTags
End Property

Private Sub pending(pending_msg)
    TFeatureRunner.pending pending_msg
End Sub

Public Sub Given_an_empty_step()

End Sub

Public Sub Given_a_valid_step()
    TSpec.expect(1).to_be 1
End Sub

Public Sub Given_an_invalid_step()
    TSpec.expect(1).to_be 2
End Sub

