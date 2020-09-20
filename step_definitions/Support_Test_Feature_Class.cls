VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Support_Test_Feature_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this class serves as a substitute for dynamically created code(step implementations)

Option Explicit

'separate tags by comma, e.g.: "wip,debug,production"
Const cTags = ""

Dim m_description As String

Public Property Get Description() As String
    
    If m_description = vbNullString Then
        m_description = "sample feature object for testing the framework"
    End If
    Description = m_description
End Property

Public Sub test_examples()

    Dim Examples As Collection
    Dim example As Variant

    Set Examples = New Collection
    
    For Each example In Examples
        TExampleRunner_Old.run_example example, Me
    Next
End Sub

Public Function run_step(step_definition As Variant) As Variant
    
    On Error GoTo error_handler
    
    Select Case step_definition.Item(STEP_ATTR_TYPE) & " " & step_definition.Item(STEP_ATTR_NAME)
        
        Case "Given a step with an empty implementation"
        
        Case "Given a step with a valid expectation"
            TSpec.expect(1).to_be 1
        
        Case "Given a step with an invalid expectation"
            TSpec.expect(1).to_be 2
        
        Case "Given a step with a valid implementation"
        
        Case "Given a step with an invalid implementation"
            Err.Raise 7777, "sample for an failed test step"
            
        Case Else
            Err.Raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array("OK")
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

Public Property Let Description(ByVal new_description As String)
    m_description = new_description
End Property
