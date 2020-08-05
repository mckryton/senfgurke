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
Dim m_feature_desc As String
Dim m_formatted_feature_desc As String
Dim m_verbose_formatter As TReport_Formatter_Verbose

Public Property Get Description() As String

    Description = "The verbose formatter will show the test results" & vbLf _
                    & "in a detailed format."
End Property

' in the future this property shall be replaced with a function that parses text
Public Property Get Examples() As Collection

    Dim examples_list As Collection
    Dim example As Variant

    Set examples_list = New Collection
    example = Array("Rule: verbose report will indent the feature description by one tab")
    examples_list.Add example
    example = Array("Example: feature description with 2 lines", _
                    "Given a feature with a description ""description line 1 <newline> description line 2""", _
                    "And the report format is set to verbose", _
                    "When the feature description is formatted", _
                    "Then every line of the description is indented by one tab")
    examples_list.Add example
    example = Array("Example: feature description with a single line", _
                "Given a feature with a description ""description line 1""", _
                "And the report format is set to verbose", _
                "When the feature description is formatted", _
                "Then every line of the description is indented by one tab")
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

    Dim desc_line As Variant
    
    On Error GoTo error_handler
    Select Case pcolStepProps.Item(TExampleRunner.ATTR_STEP_HEAD) & " " & pcolStepProps.Item(TExampleRunner.ATTR_STEP_BODY)
    
        Case "Given a feature with a description ""description line 1 <newline> description line 2"""
            m_feature_desc = "description line 1" & vbLf & "description line 2"
        
        Case "Given a feature with a description ""description line 1"""
            m_feature_desc = "description line 1"
    
        Case "Given the report format is set to verbose"
            Set m_verbose_formatter = New TReport_Formatter_Verbose
            
        Case "When the feature description is formatted"
            m_formatted_feature_desc = m_verbose_formatter.format_description(m_feature_desc)
            
        Case "Then every line of the description is indented by one tab"
            For Each desc_line In Split(m_formatted_feature_desc, vbLf)
                TSpec.expect(Left(desc_line, 1)).to_be vbTab
            Next
           
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
