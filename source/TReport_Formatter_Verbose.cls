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

Dim m_current_feature As String
Dim m_current_rule As String
Dim m_current_example As String

Public Function process_msg(msg_package As Collection, Optional silent) As String

    Dim message As String

    If IsMissing(silent) Then silent = False
    message = prepare_msg(msg_package)
    Select Case msg_package(TReport.MSG_TYPE)
        Case TReport.TYPE_FEATURE_NAME
            CurrentFeature = message & vbLf
        Case TReport.TYPE_DESC
            CurrentFeature = CurrentFeature & message & vbLf
        Case TReport.TYPE_RULE
            CurrentRule = message & vbLf
        Case TReport.TYPE_EXAMPLE_TITLE
            CurrentExample = message & vbLf
        Case TReport.TYPE_STEP
            message = Join(Split(CurrentFeature & CurrentRule & CurrentExample & message, vbLf), vbLf)
            'report feature names, rules and example titles only when example steps are reported
            If Not silent Then Debug.Print message
            process_msg = message
            CurrentFeature = vbNullString
            CurrentRule = vbNullString
            CurrentExample = vbNullString
        Case Else
            Debug.Print "DEBUG: unknown report msg type >" & msg_package(TReport.MSG_TYPE) & "<"
            process_msg = vbNullString
    End Select
End Function

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

    formatted_msg = Space$(4) & " " & status & Space$(10 - 1 - Len(status)) & step
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

Public Function format_feature_description(description As String) As String
    format_feature_description = Space$(4) & Join(Split(description, vbLf), vbLf + vbTab)
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
