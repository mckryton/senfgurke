VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this is a template for new example driven tests

Option Explicit

'separate tags by comma, e.g.: "wip,debug,production"
Const cTags = ""

Public Property Get description() As String

    description = "Add a description" & vbLf & _
                    vbTab & "to explain your features purpose."
End Property

Public Sub test_examples()

    Dim examples As Collection
    Dim example As Variant

    Set examples = New Collection
    example = Array("Rule: specify a rule to classify the following examples")
    examples.Add example
    example = Array("Example: add a headline to identify this example", _
                        "Given a precondition", _
                        "And another precondition", _
                        "When any action happens", _
                        "Then some result is expected")
    examples.Add example
    
    For Each example In examples
        TExampleRunner.run_example example, Me
    Next
End Sub

Public Function run_step(pcolStepProps As Collection) As Variant
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item("step_type") & pcolStepProps.Item("line_body")
    
        Case "Given a precondition"
           
        Case "Given another precondition"
            
        Case "When any action happens"

        Case "Then some result is expected"
        
        Case Else
            Err.raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array("OK")
    Exit Function

error_handler:
    run_step = TCase.raise_step_error(Err.Number, Err.description)
End Function

Public Sub before()
    
End Sub

Public Sub after()

End Sub

Public Property Get Tags() As Variant
    Tags = cTags
End Property

Private Sub pending(pending_msg)
    TCase.pending pending_msg
End Sub

Private Property Get Log() As logger
    
    Set Log = TCase.Log
End Property

