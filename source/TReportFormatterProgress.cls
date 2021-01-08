VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TReportFormatterProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_step_templates As Collection
Dim m_reporter As TReport

Public Function process_msg(msg_package As Collection, Optional silent) As String

    Dim message As String

    If IsMissing(silent) Then silent = False
    message = format_msg_content(msg_package)
    Select Case msg_package(REPORT_MSG_TYPE)
        Case REPORT_MSG_TYPE_FEATURE_NAME
            
        Case REPORT_MSG_TYPE_RULE
            
        Case REPORT_MSG_TYPE_EXAMPLE_TITLE
            
        Case REPORT_MSG_TYPE_STEP
            process_msg = message
            Reporter.print_report message, consecutive:=True, silent:=silent
        Case Else
            process_msg = vbNullString
    End Select
End Function

Public Function format_msg_content(msg_package As Collection) As String
    
    Dim formatted_msg As String
    
    Select Case msg_package(REPORT_MSG_TYPE)
        Case REPORT_MSG_TYPE_STEP
            formatted_msg = format_step(msg_package(REPORT_MSG_CONTENT), msg_package(REPORT_MSG_STATUS), msg_package(REPORT_MSG_ERR))
        Case REPORT_MSG_TYPE_CODE_TEMPLATE
            StepTemplates.Add msg_package(REPORT_MSG_CONTENT)
        Case REPORT_MSG_TYPE_DURATION
            Reporter.print_report vbLf & message, silent:=silent
        Case Else
            formatted_msg = vbNullString
    End Select
    format_msg_content = formatted_msg
End Function

Public Function format_step(step_name As String, status As String, err_msg As String) As String
    
    Dim formatted_msg As String
    Dim step_lines As Variant
    Dim line_index As Integer

    Select Case status
        Case STATUS_OK
            formatted_msg = "."
        Case STATUS_MISSING
            formatted_msg = "M"
        Case STATUS_PENDING
            formatted_msg = "P"
        Case STATUS_FAIL
            formatted_msg = "F"
        Case Else
            Debug.Print "unknown step status >" & status & "<"
    End Select
    'add err msg if available
    If err_msg <> vbNullString Then
        'add feature name + example name
        formatted_msg = formatted_msg & vbLf & Space$(10) & err_msg
    End If
    format_step = formatted_msg
End Function

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
        report_output = vbLf & CStr(template)
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

