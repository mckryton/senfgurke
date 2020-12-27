VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_report_formatter As Variant

Public Sub report(message_type As String, message_content As String, Optional status, Optional err_msg)

    Dim msg_package As Collection
    
    If Not Report_Formatter Is Nothing Then
        Set msg_package = New Collection
        If IsMissing(status) Then
            msg_package.Add vbNullString, REPORT_MSG_STATUS
        Else
            msg_package.Add status, REPORT_MSG_STATUS
        End If
        If IsMissing(err_msg) Then
            msg_package.Add vbNullString, REPORT_MSG_ERR
        Else
            msg_package.Add err_msg, REPORT_MSG_ERR
        End If
        Set msg_package = build_msg_package(message_type, message_content, CStr(status), CStr(err_msg))
        Report_Formatter.process_msg msg_package
        Set msg_package = Nothing
    End If
End Sub

Public Function build_msg_package(message_type As String, message_content As String, Optional status As String, Optional err_msg As String) As Collection

    Dim msg_package As Collection
    
    If IsMissing(status) Then status = vbNullString
    If IsMissing(err_msg) Then err_msg = vbNullString
    Set msg_package = New Collection
    With msg_package
        .Add message_type, REPORT_MSG_TYPE
        .Add message_content, REPORT_MSG_CONTENT
        .Add status, REPORT_MSG_STATUS
        .Add err_msg, REPORT_MSG_ERR
    End With
    Set build_msg_package = msg_package
End Function

Public Property Get Report_Formatter() As Variant
    If Not IsObject(m_report_formatter) Then
        Set m_report_formatter = New TReportFormatterVerbose
    End If
    Set Report_Formatter = m_report_formatter
End Property

Public Property Let Report_Formatter(ByVal new_report_formatter As Variant)
    Set m_report_formatter = new_report_formatter
End Property

Public Sub report_code_templates_for_missing_steps()

    Dim code_templates As String
    
    code_templates = Report_Formatter.get_code_templates_for_missing_steps
    If code_templates <> vbNullString Then
        Debug.Print
        Debug.Print
        Debug.Print code_templates
    End If
End Sub
