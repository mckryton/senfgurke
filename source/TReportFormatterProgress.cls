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

Dim m_step_implementation_templates As Collection
Dim m_reporter As TReport
Dim m_step_count As Long


Public Function process_msg(msg_package As Collection) As String

    Dim message As String

    message = format_msg_content(msg_package)
    Select Case msg_package(REPORT_MSG_TYPE)
        Case REPORT_MSG_TYPE_STEP
            count_step
            If step_counter Mod 80 = 0 And step_counter > 0 Then message = message & vbLf
            Reporter.print_report message, consecutive:=True
        Case REPORT_MSG_TYPE_CODE_TEMPLATE
            Reporter.add_code_template_to_buffer msg_package(REPORT_MSG_CONTENT)
        Case REPORT_MSG_TYPE_PARSE_ERR
            Reporter.print_report message
        Case REPORT_MSG_TYPE_STATS
            message = vbLf & message
            Reporter.print_report message
    End Select
    process_msg = message
End Function

Public Function format_msg_content(msg_package As Collection) As String
    
    Dim formatted_msg As String
    
    Select Case msg_package(REPORT_MSG_TYPE)
        Case REPORT_MSG_TYPE_STEP
            formatted_msg = format_step(msg_package(REPORT_MSG_CONTENT), msg_package(REPORT_MSG_STATUS), msg_package(REPORT_MSG_ERR))
        Case REPORT_MSG_TYPE_CODE_TEMPLATE
            Reporter.add_code_template_to_buffer msg_package(REPORT_MSG_CONTENT)
        Case REPORT_MSG_TYPE_STATS
            formatted_msg = vbLf & msg_package(REPORT_MSG_CONTENT)
        Case REPORT_MSG_TYPE_PARSE_ERR
            formatted_msg = format_parse_error(msg_package(REPORT_MSG_CONTENT), msg_package(REPORT_MSG_ERR))
        Case Else
            'msg doesn't need to be formatted
            formatted_msg = msg_package(REPORT_MSG_CONTENT)
    End Select
    format_msg_content = formatted_msg
End Function

Public Function format_parse_error(err_msg As String, err_description As String) As String
    format_parse_error = "Error: " & err_msg & vbLf _
                             & err_description
'                            & Space$(2) & err_description
End Function

Public Function format_step(step_name As String, status As String, err_msg As String) As String
    
    Dim formatted_msg As String

    Select Case status
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
            Reporter.add_err_msg_to_buffer "Err in step: " & step_name
        Case Else
            Debug.Print "unknown step status >" & status & "<"
    End Select
    'add err msg if available
    If err_msg <> vbNullString Then
        Reporter.add_err_msg_to_buffer err_msg
    End If
    format_step = formatted_msg
End Function

Private Property Get step_implementation_templates() As Collection
    
    If m_step_implementation_templates Is Nothing Then
        Set m_step_implementation_templates = New Collection
    End If
    Set step_implementation_templates = m_step_implementation_templates
End Property

Public Property Get Reporter() As TReport
    Set Reporter = m_reporter
End Property

Public Property Let Reporter(ByVal new_reporter As TReport)
    Set m_reporter = new_reporter
End Property

Private Property Get step_counter() As Variant
    'TODO: refactor get step_count from Session.Statistics
    step_counter = m_step_count
End Property

Public Sub count_step()
    m_step_count = m_step_count + 1
End Sub

Private Sub Class_Initialize()
    m_step_count = 0
End Sub

Private Sub Class_Terminate()
    Set m_step_implementation_templates = Nothing
End Sub
