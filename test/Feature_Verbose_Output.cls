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
Const cTags = "wip"

Public Property Get description() As String

    description = "The verbose formatter will show the test results" & vbLf _
                    & "in a detailed format."
End Property

Public Sub test_examples()

    Dim examples As Collection
    Dim example As Variant

    Set examples = New Collection
    example = Array("Rule: Before executing example steps the feature description will be printed on console")
    examples.Add example
    example = Array("Rule: A step will be preceded with it's result in capital letters")
    examples.Add example
    example = Array("Rule: Rules and example titles will add an empty line before them")
    examples.Add example
    example = Array("Rule: Failed steps will show the error message after the step")
    examples.Add example
    example = Array("Rule: Pending steps will show a pending message after the step")
    examples.Add example
    example = Array("Rule: Missing steps will show a code snippet after the step")
    examples.Add example

    For Each example In examples
        TExampleRunner.run_example example, Me
    Next
End Sub

Public Function run_step(pcolStepProps As Collection) As Variant
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item(TExampleRunner.ATTR_STEP_HEAD) & " " & pcolStepProps.Item(TExampleRunner.ATTR_STEP_BODY)
    
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
    run_step = TFeature.fail_step(Err.Number, Err.description)
End Function

Public Sub before_example()
    
End Sub

Public Sub after_example()

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

