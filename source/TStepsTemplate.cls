VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TStepsTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this is a template for new example driven tests

Option Explicit

'separate tags by comma, e.g.: "wip,debug,production"
Const cTags = ""
Dim m_description As String

Public Property Get description() As String
    
    If m_description = vbNullString Then
        m_description = "Add a description" & vbLf _
                            & "to explain your features purpose."
    End If
    description = m_description
End Property

' in the future this property shall be replaced with a function that parses text
Public Property Get Examples() As Collection

    Dim examples_list As Collection
    Dim example As Variant

    Set examples_list = New Collection
    example = Array("Rule: specify a rule to classify the following examples_list")
    examples_list.Add example
    example = Array("Example: add a headline to identify this example", _
                        "Given a precondition", _
                        "And another precondition", _
                        "When any action happens", _
                        "Then some result is expected")
    examples_list.Add example
    Set Examples = examples_list
End Property

Public Function run_step(step_definition As Collection) As Variant
    
    On Error GoTo error_handler
    
    Select Case step_definition.Item(ATTR_STEP_HEAD) & " " & step_definition.Item(TExampleRunner_Old.ATTR_STEP_BODY)
        
        Case "Given a precondition"
           
        Case "Given another precondition"
            
        Case "When any action happens"

        Case "Then some result is expected"
        
        Case Else
            Err.Raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array("OK")
    Exit Function

error_handler:
    run_step = TStepRunner.fail_step(Err.Number, Err.description)
End Function

Public Sub before_example()
    
End Sub

Public Sub after_example()

End Sub

Public Property Get tags() As Variant
    tags = cTags
End Property

Private Sub pending(pending_msg)
    TStepRunner.pending pending_msg
End Sub

Public Property Let description(ByVal new_description As String)
    m_description = new_description
End Property
