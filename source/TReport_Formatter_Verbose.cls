VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TReport_Formatter_Verbose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub process_msg(msg_package As Collection)

    Dim message As String

    message = prepare_msg(msg_package)
    Debug.Print message
End Sub

Public Function prepare_msg(msg_package As Collection) As String
    
    Dim formatted_msg As String
    
    Select Case msg_package(TReport.MSG_TYPE)
        Case TReport.TYPE_FEATURE_NAME
            formatted_msg = format_feature_name(msg_package(TReport.MSG_CONTENT))
        Case TReport.TYPE_DESC
            formatted_msg = format_feature_description(msg_package(TReport.MSG_CONTENT))
        Case TReport.TYPE_RULE
            formatted_msg = format_rule(msg_package(TReport.MSG_CONTENT))
        Case TReport.TYPE_EXAMPLE_TITLE
            formatted_msg = format_example_title(msg_package(TReport.MSG_CONTENT))
        Case TReport.TYPE_STEP
            formatted_msg = format_step(msg_package(TReport.MSG_CONTENT), msg_package(TReport.MSG_STATUS), msg_package(TReport.MSG_ERR))
        Case Else
            formatted_msg = vbNullString
    End Select
    prepare_msg = formatted_msg
End Function

Public Function format_step(step As String, status As String, err_msg As String) As String
    
    Dim formatted_msg As String

    formatted_msg = vbTab & " " & status & Space$(10 - 1 - Len(status)) & step
    If err_msg <> vbNullString Then
        formatted_msg = formatted_msg & vbLf & Space$(10) & err_msg
    End If
    format_step = formatted_msg
End Function

Public Function format_example_title(example_title As String) As String
    format_example_title = vbLf & vbTab & example_title
End Function

Public Function format_feature_name(feature_name As String) As String
    format_feature_name = vbLf & "Test case: " & feature_name
End Function

Public Function format_feature_description(Description As String) As String
    format_feature_description = vbTab & Join(Split(Description, vbLf), vbLf + vbTab)
End Function

Public Function format_rule(rule As String) As String
    format_rule = vbLf & Join(Split(rule, vbLf), vbLf + Space$(7))
End Function
