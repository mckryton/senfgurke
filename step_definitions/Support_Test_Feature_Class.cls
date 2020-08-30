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

Public Property Get description() As String
    
    If m_description = vbNullString Then
        m_description = "sample feature object for testing the framework"
    End If
    description = m_description
End Property

Public Sub test_examples()

    Dim Examples As Collection
    Dim example As Variant

    Set Examples = New Collection
    
    For Each example In Examples
        TExampleRunner.run_example example, Me
    Next
End Sub

Public Function run_step(pcolStepProps As Collection) As Variant
    
    On Error GoTo error_handler
    
    Select Case pcolStepProps.Item(TExampleRunner.ATTR_STEP_HEAD) & " " & pcolStepProps.Item(TExampleRunner.ATTR_STEP_BODY)
        
        Case "Given a step with an empty implementation"
        
        Case "Given a step with a valid expectation"
            TSpec.expect(1).to_be 1
        
        Case "Given a step with an invalid expectation"
            TSpec.expect(1).to_be 2
        
        Case Else
            Err.Raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array("OK")
    Exit Function

error_handler:
    run_step = TFeature.fail_step(Err.Number, Err.description)
End Function

Public Sub before()
    
End Sub

Public Sub after()

End Sub

Public Property Get tags() As Variant
    tags = cTags
End Property

Private Sub pending(pending_msg)
    TFeature.pending pending_msg
End Sub

Public Property Let description(ByVal new_description As String)
    m_description = new_description
End Property
