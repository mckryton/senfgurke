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
Dim m_test_feature_object As Support_Test_Feature_Class

Public Property Get Description() As String

    Description = "The verbose formatter will show the test results" & vbLf _
                    & "in a detailed format."
End Property

' in the future this property shall be replaced with a function that parses text
Public Property Get Examples() As Collection

    Dim examples_list As Collection
    Dim example As Variant

    Set examples_list = New Collection
    example = Array("Rule: verbose report will insert a tab after every linebreak in the dscription")
    examples_list.Add example
    example = Array("Example: feature with description", _
                    "Given a feature with a description ""this is the description of a feature""", _
                    "When the feature is executed", _
                    "Then the description is printed on the console with an extra tab inserted on the start of the second line")
    examples_list.Add example
    example = Array("Rule: A step will be preceded with it's result in capital letters")
    examples_list.Add example
    example = Array("Rule: Rules and example titles will add an empty line before them")
    examples_list.Add example
    example = Array("Rule: Failed steps will show the error message after the step")
    examples_list.Add example
    example = Array("Rule: Pending steps will show a pending message after the step")
    examples_list.Add example
    example = Array("Rule: Missing steps will show a code snippet after the step")
    examples_list.Add example
    Set Examples = examples_list
End Property

Public Function run_step(pcolStepProps As Collection) As Variant
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item(TExampleRunner.ATTR_STEP_HEAD) & " " & pcolStepProps.Item(TExampleRunner.ATTR_STEP_BODY)
    
        Case "Given a feature with a description ""this is the description <newline> of a feature"""
            Set m_test_feature_object = New Support_Test_Feature_Class
            m_test_feature_object.Description = "this is the description of a feature"
           
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

