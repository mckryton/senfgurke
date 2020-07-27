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

    Select Case msg_package(TReport.MSG_TYPE)
        Case TReport.TYPE_FEATURE_NAME
            print_feature_name msg_package(TReport.MSG_CONTENT)
        Case TReport.TYPE_DESC
            print_description msg_package(TReport.MSG_CONTENT)
        Case TReport.TYPE_RULE
            print_rule msg_package(TReport.MSG_CONTENT)
        Case TReport.TYPE_EXAMPLE_TITLE
            print_example_title msg_package(TReport.MSG_CONTENT)
        Case TReport.TYPE_STEP
            print_step msg_package(TReport.MSG_CONTENT), msg_package(TReport.MSG_STATUS), msg_package(TReport.MSG_ERR)
    End Select
End Sub

Public Sub print_example_title(example_title As String)
    Debug.Print format_example_title(example_title)
End Sub

Public Sub print_feature_name(feature_name As String)
    Debug.Print format_feature_name(feature_name)
End Sub

Public Sub print_description(Description As String)
    Debug.Print format_description(Description)
End Sub

Public Sub print_rule(rule As String)
    Debug.Print format_rule(rule)
End Sub

Public Sub print_step(step As String, status As String, err_msg As String)
    Debug.Print vbTab & status, step
    If err_msg <> vbNullString Then
        Debug.Print "", err_msg
    End If
End Sub

Public Function format_example_title(example_title As String) As String
    format_example_title = vbLf & vbTab & example_title
End Function

Public Function format_feature_name(feature_name As String) As String
    format_feature_name = vbLf & "Test case: " & feature_name
End Function

Public Function format_description(Description As String) As String
    format_description = vbTab & Join(Split(Description, vbLf), vbLf + vbTab)
End Function

Public Function format_rule(rule As String) As String
    format_rule = vbLf & vbTab & rule
End Function
