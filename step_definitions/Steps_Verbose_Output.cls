VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_Verbose_Output"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'separate tags by comma, e.g.: "wip,debug,production"
Const cTags = ""
Dim m_report_output As String
Dim m_report_msg As Collection
Dim m_verbose_formatter As TReport_Formatter_Verbose

Public Property Get Description() As String

    Description = "While executing examples senfgurke will send messages about progress" & vbLf _
                    & "and success for reporting. The verbose formatter will turn those messages into" & vbLf _
                    & " a verbose report that is printed on the debug console."
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
                "When the reported message is prepared as output for the report", _
                "Then every line of the description is indented by one tab")
    examples_list.Add example
    example = Array("Example: feature description with a single line", _
                "Given a feature with a description ""description line 1""", _
                "And the report format is set to verbose", _
                "When the reported message is prepared as output for the report", _
                "Then every line of the description is indented by one tab")
    examples_list.Add example
    
    example = Array("Rule: A step will be preceded with it's result in capital letters initially intendented by 1 tab " & vbLf _
                    & "and then indented by 1 space and steps following after indented by 10 spaces")
    examples_list.Add example
    example = Array("Example: successful step", _
                "Given a step ""this is a step"" with the status ""OK""", _
                "And the report format is set to verbose", _
                "When the reported message is prepared as output for the report", _
                "Then the resulting output is set to tab and "" OK       this is a step""")
    examples_list.Add example
    example = Array("Example: failed step", _
                "Given a step ""this is a step"" with the status ""FAILED""", _
                "And the report format is set to verbose", _
                "When the reported message is prepared as output for the report", _
                "Then the resulting output is set to tab and "" FAILED   this is a step""")
    examples_list.Add example
    example = Array("Example: missing step", _
                "Given a step ""this is a step"" with the status ""MISSING""", _
                "And the report format is set to verbose", _
                "When the reported message is prepared as output for the report", _
                "Then the resulting output is set to tab and "" MISSING  this is a step""")
    examples_list.Add example
    
    example = Array("Rule: Rules and example titles will add an empty line before them " & vbLf _
                    & "and indent following lines with 3 tabs")
    examples_list.Add example
    example = Array("Example: Rule over two lines", _
                "Given a rule ""this is a sample rule<br>and this is line 2 of the same rule""", _
                "And the report format is set to verbose", _
                "When the reported message is prepared as output for the report", _
                "Then the resulting output starts with a new line", _
                "And every line after the first one is indented by 7 spaces")
    examples_list.Add example
    
    example = Array("Rule: Failed steps will show the error message after the step indented")
    examples_list.Add example
    example = Array("Example: Step fails with an error message", _
                "Given a step ""this step fails"" with status ""FAILED"" and with the error message ""err: sample err msg""", _
                "And the report format is set to verbose", _
                "When the reported message is prepared as output for the report", _
                "Then the resulting output starts with the step status and description", _
                "And the second line shows the indented error message")
    examples_list.Add example
    
'    example = Array("Rule: Pending steps will show a pending message after the step")
'    examples_list.Add example
'    example = Array("Rule: Missing steps will show a code snippet after the step")
'    examples_list.Add example
    Set Examples = examples_list
End Property

Public Function run_step(step_definition As Collection) As Variant

    Dim output_line As Variant
    Dim output_lines As Variant
    Dim index As Integer
    
    On Error GoTo error_handler
    Select Case step_definition.Item(TExampleRunner_Old.ATTR_STEP_HEAD) & " " & step_definition.Item(TExampleRunner_Old.ATTR_STEP_BODY)
    
        Case "Given a feature with a description ""description line 1 <newline> description line 2"""
            Set m_report_msg = TReport.build_msg_package(TReport.TYPE_DESC, "description line 1" & vbLf & "description line 2", vbNullString, vbNullString)
        
        Case "Given a feature with a description ""description line 1"""
            Set m_report_msg = TReport.build_msg_package(TReport.TYPE_DESC, "description line 1", vbNullString, vbNullString)

        Case "Given the report format is set to verbose"
            Set m_verbose_formatter = New TReport_Formatter_Verbose
            
        Case "Given a step ""this is a step"" with the status ""OK"""
            Set m_report_msg = TReport.build_msg_package(TReport.TYPE_STEP, "this is a step", "OK", vbNullString)
            
        Case "Given a step ""this is a step"" with the status ""FAILED"""
            Set m_report_msg = TReport.build_msg_package(TReport.TYPE_STEP, "this is a step", "FAILED", vbNullString)
        
        Case "Given a step ""this is a step"" with the status ""MISSING"""
            Set m_report_msg = TReport.build_msg_package(TReport.TYPE_STEP, "this is a step", "MISSING", vbNullString)
        
        Case "Given a rule ""this is a sample rule<br>and this is line 2 of the same rule"""
             Set m_report_msg = TReport.build_msg_package(TReport.TYPE_RULE, "this is a sample rule" & vbLf & "and this is line 2 of the same rule", vbNullString, vbNullString)
             
        Case "Given a step ""this step fails"" with status ""FAILED"" and with the error message ""err: sample err msg"""
            Set m_report_msg = TReport.build_msg_package(TReport.TYPE_STEP, "this step fails", "FAILED", "err: sample err msg")
            
        Case "When the reported message is prepared as output for the report"
            m_report_output = m_verbose_formatter.prepare_msg(m_report_msg)
            
        Case "Then every line of the description is indented by one tab"
            For Each output_line In Split(m_report_output, vbLf)
                TSpec.expect(Left(output_line, 1)).to_be vbTab
            Next
    
        Case "Then the resulting output is set to tab and "" OK       this is a step"""
            TSpec.expect(m_report_output).to_be vbTab & " OK       this is a step"
        
        Case "Then the resulting output is set to tab and "" FAILED   this is a step"""
            TSpec.expect(m_report_output).to_be vbTab & " FAILED   this is a step"
        
        Case "Then the resulting output is set to tab and "" MISSING  this is a step"""
            TSpec.expect(m_report_output).to_be vbTab & " MISSING  this is a step"
            
        Case "Then the resulting output starts with a new line"
            TSpec.expect(Left(m_report_output, 1)).to_be vbLf
            
        Case "Then every line after the first one is indented by 7 spaces"
            output_lines = Split(m_report_output, vbLf)
            For index = 2 To UBound(output_lines)
                TSpec.expect(Left(output_lines(index), 7)).to_be Space$(7)
            Next
            
        Case "Then the resulting output starts with the step status and description"
            TSpec.expect(Split(m_report_output, vbLf)(0)).to_be vbTab & " FAILED   this step fails"
        
        Case "Then the second line shows the indented error message"
            TSpec.expect(Split(m_report_output, vbLf)(1)).to_be Space$(10) & "err: sample err msg"
           
        Case Else
            Err.Raise ERR_ID_STEP_IS_MISSING
    End Select
    run_step = Array("OK")
    Exit Function

error_handler:
    run_step = TExampleRunner.fail_step(Err.Number, Err.Description)
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
