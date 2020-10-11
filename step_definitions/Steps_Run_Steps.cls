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

Dim m_example_step As TStep
Dim m_step_result As Variant
Dim m_step_implementation_name As String
Dim m_step_implementation_template As String

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
            
        Case "Given a step ""Given a valid step"""
            Set m_example_step = New TStep
            With m_example_step
                .SType = STEP_TYPE_GIVEN
                .Name = "a valid step"
            End With
        
        Case "When the step is executed"
            m_step_result = TStepRunner.run_step(m_example_step)
        
        Case "When the step is translated into a function name"
            m_step_implementation_name = m_example_step.get_step_function_name()
            
        Case "When the code template for the step implementation is requested"
            m_step_implementation_template = m_example_step.get_step_function_template
        
        Case "Then the execution result is ""OK"""
            TSpec.expect(m_step_result(0)).to_be "OK"
            
        Case "Then the execution result is ""FAIL"""
            TSpec.expect(m_step_result(0)).to_be "FAIL"
            
        Case "Then the execution result is ""MISSING"""
            TSpec.expect(m_step_result(0)).to_be "MISSING"
        
        Case "Then the name of the resulting function is ""Given_a_step"""
            TSpec.expect(Left(m_step_implementation_name, Len("Given_a_step"))).to_be "Given_a_step"
            TSpec.expect(Len(m_step_implementation_name)).to_be Len("Given_a_step_xxXXxxXXxxXX")
            
        Case "Then the code template for the step implementation is " & _
                            """Public Sub Given_a_valid_step_8A74152FD2F9()<br>" _
                            & "    'Given a valid step<br><br>" _
                            & "End Sub"""
            TSpec.expect(m_step_implementation_template).to_be "Public Sub Given_a_valid_step_8A74152FD2F9()" & vbLf _
                                                                & "    'Given a valid step" & vbLf _
                                                                & vbLf _
                                                                & "End Sub"
        
        Case Else
            Err.Raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array(STEP_RESULT_OK)
    Exit Function

error_handler:
    run_step = TStepRunner.fail_step(Err.Number, Err.description)
End Function


Public Sub Given_an_empty_step_8A7408D743D5()

End Sub

Public Sub Given_a_valid_step_8A74152FD2F9()
    TSpec.expect(1).to_be 1
End Sub

Public Sub Given_an_invalid_step_CF4B201BCF1B()
    TSpec.expect(1).to_be 2
End Sub

