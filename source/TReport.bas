Attribute VB_Name = "TReport"
Option Explicit

Public Const MSG_TYPE = "message_type"
Public Const MSG_CONTENT = "message_content"
Public Const MSG_STATUS = "status"
Public Const MSG_ERR = "error_message"

Public Const TYPE_FEATURE_NAME = "feature_name"
Public Const TYPE_DESC = "description"
Public Const TYPE_RULE = "rule"
Public Const TYPE_EXAMPLE_TITLE = "example_title"
Public Const TYPE_STEP = "step"

Dim m_report_formatter As Variant

Public Sub report(message_type As String, message_content As String, Optional status, Optional err_msg)

    Dim msg_package As Collection
    
    If Not Report_Formatter Is Nothing Then
        Set msg_package = New Collection
        If IsMissing(status) Then
            msg_package.Add vbNull, TReport.MSG_STATUS
        Else
            msg_package.Add status, TReport.MSG_STATUS
        End If
        If IsMissing(err_msg) Then
            msg_package.Add vbNull, TReport.MSG_ERR
        Else
            msg_package.Add err_msg, TReport.MSG_ERR
        End If
        Set msg_package = build_msg_package(message_type, message_content, CStr(status), CStr(err_msg))
        Report_Formatter.process_msg msg_package
        Set msg_package = Nothing
    End If
End Sub

Public Function build_msg_package(message_type As String, message_content As String, status As String, err_msg As String) As Collection

    Dim msg_package As Collection
    
    Set msg_package = New Collection
    With msg_package
        .Add message_type, TReport.MSG_TYPE
        .Add message_content, TReport.MSG_CONTENT
        .Add status, TReport.MSG_STATUS
        .Add err_msg, TReport.MSG_ERR
    End With
    Set build_msg_package = msg_package
End Function

Public Property Get Report_Formatter() As Variant
    If Not IsObject(m_report_formatter) Then
        Set m_report_formatter = New TReport_Formatter_Verbose
    End If
    Set Reporter = m_report_formatter
End Property

Public Property Let Report_Formatter(ByVal new_report_formatter As Variant)
    Set m_report_formatter = new_reporter
End Property


