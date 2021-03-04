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
Dim m_step_templates As Collection
Dim m_reporter As TReport

Public Function process_msg(msg_package As Collection, Optional silent) As String

    Dim message As String

    If IsMissing(silent) Then silent = False
    message = format_msg_content(msg_package)
    Select Case msg_package(REPORT_MSG_TYPE)
        Case REPORT_MSG_TYPE_FEATURE_NAME
            CurrentFeature = message & vbLf
        Case REPORT_MSG_TYPE_DESC
            CurrentFeature = CurrentFeature & message & vbLf
        Case REPORT_MSG_TYPE_RULE
            CurrentRule = message & vbLf
        Case REPORT_MSG_TYPE_EXAMPLE_TITLE
            CurrentExample = message & vbLf
        Case REPORT_MSG_TYPE_STEP
            message = Join(Split(CurrentFeature & CurrentRule & CurrentExample & message, vbLf), vbLf)
            'report feature names, rules and example titles only when example steps are reported
            Reporter.print_report message, silent:=silent
            process_msg = message
            CurrentFeature = vbNullString
            CurrentRule = vbNullString
            CurrentExample = vbNullString
        Case REPORT_MSG_TYPE_CODE_TEMPLATE
            StepTemplates.Add message
        Case REPORT_MSG_TYPE_STATS
            Reporter.print_report message, silent:=silent
        Case Else
            TRun.Session.Log.debug_log "unknown report msg type >" & msg_package(REPORT_MSG_TYPE) & "<"
            process_msg = vbNullString
    End Select
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
        Case REPORT_MSG_TYPE_STEP
            formatted_msg = format_step(msg_package(REPORT_MSG_CONTENT), msg_package(REPORT_MSG_STATUS), msg_package(REPORT_MSG_ERR))
        Case REPORT_MSG_TYPE_CODE_TEMPLATE
            formatted_msg = msg_package(REPORT_MSG_CONTENT)
        Case REPORT_MSG_TYPE_STATS
            formatted_msg = vbLf & msg_package(REPORT_MSG_CONTENT)
        Case Else
            formatted_msg = vbNullString
    End Select
    format_msg_content = formatted_msg
End Function

Public Function format_step(step_name As String, status As String, err_msg As String) As String
    
    Dim formatted_msg As String
    Dim step_lines As Variant
    Dim line_index As Integer

    step_lines = Split(step_name, vbLf)
    formatted_msg = Space$(4) & " " & status & Space$(10 - 1 - Len(status)) & step_lines(0)
    'format docstring output if available
    If UBound(step_lines) > 0 Then
        formatted_msg = formatted_msg & vbLf & Space$(16) & Trim(step_lines(1))
        For line_index = 2 To UBound(step_lines) - 1
            formatted_msg = formatted_msg & vbLf & Space$(18) & Trim(step_lines(line_index))
        Next
        formatted_msg = formatted_msg & vbLf & Space$(16) & Trim(step_lines(UBound(step_lines)))
    End If
    'add err msg if available
    If err_msg <> vbNullString Then
        formatted_msg = formatted_msg & vbLf & Space$(10) & err_msg
    End If
    format_step = formatted_msg
End Function

Public Function format_example_title(example_title As String) As String
    format_example_title = vbLf & Space$(4) & example_title
End Function

Public Function format_feature_name(feature_name As String) As String
    format_feature_name = vbLf & "Feature: " & feature_name
End Function

Public Function format_feature_description(Description As String) As String

    Dim desc_lines As Variant
    Dim line As Variant
    Dim formatted_desc As String
    
    formatted_desc = vbNullString
    desc_lines = Split(Description, vbLf)
    For Each line In desc_lines
        formatted_desc = format_feature_description & Space$(4) & Trim(line) & vbLf
    Next
    format_feature_description = Left(formatted_desc, Len(formatted_desc) - 1)
End Function

Public Function format_rule(rule As String) As String
    format_rule = vbLf & Space$(2) & Join(Split(rule, vbLf), vbLf + Space$(7))
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

Private Property Get StepTemplates() As Collection
    
    If m_step_templates Is Nothing Then
        Set m_step_templates = New Collection
    End If
    Set StepTemplates = m_step_templates
End Property

Private Sub Class_Terminate()
    Set m_step_templates = Nothing
End Sub

Public Function get_code_templates_for_missing_steps() As String

    Dim report_output As String
    Dim template As Variant
    
    report_output = vbNullString
    For Each template In StepTemplates
        report_output = report_output & vbLf & CStr(template) & vbLf
    Next
    If report_output <> vbNullString Then
        report_output = "You can implement step definitions for undefined steps with these snippets:" _
                            & vbLf & report_output
    End If
    get_code_templates_for_missing_steps = report_output
End Function

Public Property Get Reporter() As TReport
    Set Reporter = m_reporter
End Property

Public Property Let Reporter(ByVal new_reporter As TReport)
    Set m_reporter = new_reporter
End Property

