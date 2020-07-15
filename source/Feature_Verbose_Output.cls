VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feature_Verbose_Output"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'separate tags by comma, e.g.: "wip,debug,production"
Const cTags = ""

Public Property Get description() As String

    description = "Senfgurke will execute example steps" & vbLf & _
                    vbTab & "and show the result on the console."
End Property

Public Sub test_examples()

    Dim examples As Collection
    Dim example As Variant

    Set examples = New Collection
    example = Array("Rule: Before executing example steps the feature description will be printed on console")
    examples.Add example
    example = Array("Rule: Steps not starting with Given, When, Then or And will produces a Gherkin syntax error")
    examples.Add example
    example = Array("Rule: Every step will be preceded with it's result in capital letters")
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

