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

Public Function process_msg(msg_package As Collection, Optional silent) As String

    Dim message As String

    If IsMissing(silent) Then silent = False
    message = format_msg_content(msg_package)
    Select Case msg_package(TReport.MSG_TYPE)
        Case TReport.TYPE_FEATURE_NAME
            
        Case TReport.TYPE_RULE
            
        Case TReport.TYPE_EXAMPLE_TITLE
            
        Case TReport.TYPE_STEP
            process_msg = message
            If Not silent Then Debug.Print message;
        Case Else
            process_msg = vbNullString
    End Select
End Function

Public Function format_msg_content(msg_package As Collection) As String
    
    Dim formatted_msg As String
    
    Select Case msg_package(TReport.MSG_TYPE)
'        Case TReport.TYPE_FEATURE_NAME
'            formatted_msg = format_feature_name(msg_package(TReport.MSG_CONTENT))
'        Case TReport.TYPE_DESC
'            formatted_msg = format_feature_description(msg_package(TReport.MSG_CONTENT))
'        Case TReport.TYPE_RULE
'            formatted_msg = format_rule(msg_package(TReport.MSG_CONTENT))
'        Case TReport.TYPE_EXAMPLE_TITLE
'            formatted_msg = format_example_title(msg_package(TReport.MSG_CONTENT))
        Case TReport.TYPE_STEP
            formatted_msg = format_step(msg_package(TReport.MSG_CONTENT), msg_package(TReport.MSG_STATUS), msg_package(TReport.MSG_ERR))
        Case TReport.TYPE_CODE_TEMPLATE
            StepTemplates.Add msg_package(TReport.MSG_CONTENT)
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
