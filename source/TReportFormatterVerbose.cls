VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TReportFormatterVerbose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_current_feature As String
Dim m_current_rule As String
Dim m_current_example As String
Dim m_step_implementation_templates As Collection
Dim m_reporter As TReport

Private StepCounter As Long
Private IsFirstOutlineRow As Boolean
Public OutlineWasStarted As Boolean

Private Sub Class_Initialize()
    IsFirstOutlineRow = True
    OutlineWasStarted = False
    StepCounter = 0
End Sub

Private Sub Class_Terminate()
    Set m_step_implementation_templates = Nothing
End Sub
'TODO validate if this function has to return something
Public Function process_msg(msg_package As Collection) As String
    Dim formatted_msg As String
    Dim single_char_msg As String

    formatted_msg = format_msg_content(msg_package)
    process_msg = vbNullString
    Select Case msg_package(REPORT_MSG_TYPE)
        Case REPORT_MSG_TYPE_FEATURE_NAME
            CurrentFeature = formatted_msg & vbLf
        Case REPORT_MSG_TYPE_RULE
            CurrentRule = formatted_msg & vbLf
        Case REPORT_MSG_TYPE_EXAMPLE_TITLE
            CurrentExample = formatted_msg & vbLf
        Case REPORT_MSG_TYPE_STEP
            If Not OutlineWasStarted Then
                process_msg = report_step(formatted_msg)
            Else
                'report step results in single char format for each outline table row
                ' | column  1 | column 2 |
                ' ...FSS
                single_char_msg = format_step_as_progress(msg_package(REPORT_MSG_CONTENT), msg_package(REPORT_MSG_STATUS), _
                                                        msg_package(REPORT_MSG_ERR))
                report_step_as_progress single_char_msg
            End If
        Case REPORT_MSG_TYPE_OUTLINE_STEP
            process_msg = report_step(formatted_msg)
        Case REPORT_MSG_TYPE_CODE_TEMPLATE
            reporter.add_code_template_to_buffer formatted_msg
        Case REPORT_MSG_TYPE_OUTLINE_ROW
            StepCounter = 0
            reporter.print_report formatted_msg
        Case REPORT_MSG_TYPE_STATS, REPORT_MSG_TYPE_PARSE_ERR, REPORT_MSG_TYPE_OUTLINE_TABLEHEADER, REPORT_MSG_TYPE_OUTLINE_ROW
            reporter.print_report formatted_msg
    End Select
End Function

Private Function report_step(formatted_msg As String) As String
    formatted_msg = Join(Split(CurrentFeature & CurrentRule & CurrentExample & formatted_msg, vbLf), vbLf)
    If CurrentFeature = vbNullString And (CurrentRule <> vbNullString Or CurrentExample <> vbNullString) Then formatted_msg = vbLf & formatted_msg
    'report feature names, rules and example titles only when example steps are reported
    reporter.print_report formatted_msg
    CurrentFeature = vbNullString
    CurrentRule = vbNullString
    CurrentExample = vbNullString
    report_step = formatted_msg
End Function

Private Function report_step_as_progress(single_char_msg As String)
    Dim step_msg As String
    
    If StepCounter = 0 Then
        step_msg = Space$(16)
    End If
    step_msg = step_msg & single_char_msg
    StepCounter = StepCounter + 1
    reporter.print_report step_msg, consecutive:=True
End Function

Public Function format_msg_content(msg_package As Collection) As String
    Dim formatted_msg As String
    
    Select Case msg_package(REPORT_MSG_TYPE)
        Case REPORT_MSG_TYPE_FEATURE_NAME
            formatted_msg = format_feature_name(msg_package(REPORT_MSG_CONTENT))
        Case REPORT_MSG_TYPE_DESC
            formatted_msg = format_feature_description(msg_package(REPORT_MSG_CONTENT))
        Case REPORT_MSG_TYPE_RULE
            formatted_msg = format_rule(msg_package(REPORT_MSG_CONTENT))
        Case REPORT_MSG_TYPE_EXAMPLE_TITLE
            formatted_msg = format_example_title(msg_package(REPORT_MSG_CONTENT))
        Case REPORT_MSG_TYPE_PARSE_ERR
            formatted_msg = format_parse_error(msg_package(REPORT_MSG_CONTENT), msg_package(REPORT_MSG_ERR))
        Case REPORT_MSG_TYPE_STEP, REPORT_MSG_TYPE_OUTLINE_STEP
            formatted_msg = format_step(msg_package(REPORT_MSG_CONTENT), msg_package(REPORT_MSG_STATUS), msg_package(REPORT_MSG_ERR))
        Case REPORT_MSG_TYPE_STATS
            formatted_msg = vbLf & msg_package(REPORT_MSG_CONTENT)
        Case REPORT_MSG_TYPE_OUTLINE_TABLEHEADER
            formatted_msg = format_outline_table(msg_package(REPORT_MSG_CONTENT))
        Case REPORT_MSG_TYPE_OUTLINE_ROW
            formatted_msg = format_outline_row(msg_package(REPORT_MSG_CONTENT))
        Case Else
            'msg doesn't need to be formatted
            formatted_msg = msg_package(REPORT_MSG_CONTENT)
    End Select
    format_msg_content = formatted_msg
End Function

Public Function format_step(step_name As String, Status As String, err_msg As String) As String
    Dim formatted_msg As String
    Dim step_lines As Variant
    Dim line_index As Integer
    Dim is_docstring As Boolean
    Dim indention As Integer
    Dim line As String

    is_docstring = False
    step_lines = Split(step_name, vbLf)
    If Me.OutlineWasStarted Then
        'omit status on steps for outline (because steps can be repeated multiple times with different results)
        formatted_msg = Space$(14) & step_lines(0)
    Else
        formatted_msg = Space$(4) & " " & Status & Space$(10 - 1 - Len(Status)) & step_lines(0)
    End If
    'format docstring and or data table output if available
    For line_index = 1 To UBound(step_lines)
        line = Trim(step_lines(line_index))
        indention = 16
        If is_docstring Then indention = 18
        If line = """""""" Then
            is_docstring = Not is_docstring
            indention = 16
        End If
        formatted_msg = formatted_msg & vbLf & Space$(indention) & line
    Next
    'add err msg if available
    If err_msg <> vbNullString Then
        formatted_msg = formatted_msg & vbLf & Space$(10) & err_msg
    End If
    format_step = formatted_msg
End Function

Public Function format_step_as_progress(step_name As String, Status As String, err_msg As String) As String
    Dim formatted_msg As String

    Select Case Status
        Case STATUS_OK
            formatted_msg = "."
        Case STATUS_MISSING
            formatted_msg = "M"
        Case STATUS_PENDING
            formatted_msg = "P"
        Case STATUS_SKIPPED
            formatted_msg = "S"
        Case STATUS_FAIL
            formatted_msg = "F"
            reporter.add_err_msg_to_buffer "Err in step: " & step_name
        Case Else
            Debug.Print "unknown step status >" & Status & "<"
    End Select
    'add err msg if available
    If err_msg <> vbNullString Then
        reporter.add_err_msg_to_buffer Space(2) & err_msg
    End If
    format_step_as_progress = formatted_msg
End Function

Public Function format_example_title(example_title As String) As String
    format_example_title = Space$(4) & example_title
End Function

Public Function format_feature_name(feature_name As String) As String
    format_feature_name = vbLf & "Feature: " & feature_name
End Function

Public Function format_parse_error(err_msg As String, err_description As String) As String
    format_parse_error = "Error: " & err_msg & vbLf _
                            & Space$(2) & err_description
End Function

Public Function format_feature_description(description As String) As String
    Dim desc_lines As Variant
    Dim line As Variant
    Dim formatted_desc As String
    
    formatted_desc = vbNullString
    desc_lines = Split(description, vbLf)
    For Each line In desc_lines
        If Len(formatted_desc) > 0 Then formatted_desc = formatted_desc & vbLf
        formatted_desc = formatted_desc & Space$(4) & Trim(line)
    Next
    format_feature_description = formatted_desc
End Function

Public Function format_rule(rule As String) As String
    format_rule = Space$(2) & Join(Split(rule, vbLf), vbLf + Space$(7))
End Function

Public Function format_outline_table(table_header As String) As String
    Dim header_items As Variant

    IsFirstOutlineRow = True
    header_items = Split(table_header, vbLf)
    format_outline_table = vbLf & Space$(14) & header_items(0) & vbLf _
                            & Space$(16) & header_items(1)
End Function

Public Function format_outline_row(outline_row As String) As String
    If IsFirstOutlineRow Then
        format_outline_row = Space$(16) & outline_row
        IsFirstOutlineRow = False
    Else
        format_outline_row = vbLf & Space$(16) & outline_row
    End If
End Function

Private Property Get CurrentFeature() As String
    CurrentFeature = m_current_feature
End Property

Private Property Let CurrentFeature(ByVal current_feature As String)
    m_current_feature = current_feature
End Property

Private Property Get CurrentRule() As String
    CurrentRule = m_current_rule
End Property

Private Property Let CurrentRule(ByVal current_rule As String)
    m_current_rule = current_rule
End Property

Private Property Get CurrentExample() As String
    CurrentExample = m_current_example
End Property

Private Property Let CurrentExample(ByVal current_example As String)
    m_current_example = current_example
End Property

Private Property Get step_implementation_templates() As Collection
    
    If m_step_implementation_templates Is Nothing Then
        Set m_step_implementation_templates = New Collection
    End If
    Set step_implementation_templates = m_step_implementation_templates
End Property

Public Property Get reporter() As TReport
    Set reporter = m_reporter
End Property

Public Property Let reporter(ByVal new_reporter As TReport)
    Set m_reporter = new_reporter
End Property

