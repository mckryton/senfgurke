Attribute VB_Name = "TError"
Option Explicit

Private m_last_fail_msg As String

' This module sets a unique id for each error that could occur in Senfgurke.
' Unfortunately VBA handles errors differently when raised in a class in
' comparison to errors raised in a module (e.g. err.description get overwritten when
' the error was raised from a class). Please see
'  https://stackoverflow.com/questions/69370099/what-numbers-should-be-used-with-vbobjecterror
' for some background information.
' It seems that raising an error from a module with an id between 1000 and 30000
' will preserve id and description.

Public Const ERR_ID_FEATURE_SYNTAX_MISSING_DEFINITION& = 6010
Public Const ERR_ID_FEATURE_SYNTAX_MISSING_FEATURE_KEYWORD& = 6012
Public Const ERR_ID_FEATURE_SYNTAX_MISSING_FEATURE_NAME& = 6014
Public Const ERR_ID_FEATURE_SYNTAX_UNEXPECTED_LINE& = 6016

Public Const ERR_ID_SECTION_SYNTAX_DESCRIPTION_AFTER_STEPS& = 6110
Public Const ERR_ID_SECTION_SYNTAX_STEPS_IN_RULE& = 6112

Public Const ERR_ID_STEP_SYNTAX_INCOMPLETE_TABLE_ROW& = 6310
Public Const ERR_ID_STEP_SYNTAX_TABLE_WITHOUT_STEP& = 6312
Public Const ERR_ID_STEP_SYNTAX_TABLE_DUPLICATE_COLUMN& = 6314
Public Const ERR_ID_STEP_SYNTAX_TABLE_COLUMN_COUNT_MISMATCH& = 6316


Public Const ERR_ID_STEP_IS_STATUS_PENDING& = 7010
Public Const ERR_ID_STEP_IS_STATUS_MISSING& = 7020
Public Const ERR_ID_EXPECTATION_STATUS_FAILED& = 8010
Public Const ERR_ID_UNSUPPORTED_ARRAY_ERROR& = 8110
Public Const ERR_ID_UNKNOWN_MSG_TYPE& = 9010


Private Const ERR_MSG_LIST = _
                "|#" & ERR_ID_FEATURE_SYNTAX_MISSING_DEFINITION & "|" & "can't find feature definition" & _
                "|#" & ERR_ID_FEATURE_SYNTAX_UNEXPECTED_LINE & "|" & "unexpected text at line #{1}: >#{2}<" & _
                "|#" & ERR_ID_FEATURE_SYNTAX_MISSING_FEATURE_KEYWORD & "|" & "feature lacks Feature keyword at the beginning" & _
                "|#" & ERR_ID_FEATURE_SYNTAX_MISSING_FEATURE_NAME & "|" & "Feature file lacks a feature name" & _
                "|#" & ERR_ID_SECTION_SYNTAX_DESCRIPTION_AFTER_STEPS & "|" & "unexpected description after steps at line #{1}: >#{2}<" & _
                "|#" & ERR_ID_SECTION_SYNTAX_STEPS_IN_RULE & "|" & "rules are not allowed to have steps: found step at line #{1}: >#{2}<" & _
                "|#" & ERR_ID_STEP_SYNTAX_INCOMPLETE_TABLE_ROW & "|" & "closing #{pipe} is missing in table row >#{1}<" & _
                "|#" & ERR_ID_STEP_SYNTAX_TABLE_WITHOUT_STEP & "|" & "preceding step for data table is missing" & _
                "|#" & ERR_ID_STEP_SYNTAX_TABLE_DUPLICATE_COLUMN & "|" & "duplicate column names found in data table: >#{1}<" & _
                "|#" & ERR_ID_STEP_SYNTAX_TABLE_COLUMN_COUNT_MISMATCH & "|" & "column count in table row doesn't match table header >#{1}<" & _
                "|"

Public Sub raise_pre_defined_error(err_id As Long, function_name As String, Optional arguments)

    Dim description As String
    
    description = get_err_msg(err_id, arguments)
    LastFailMsg = description
    Err.raise err_id, function_name, description
End Sub

Private Function get_err_msg(err_id As Long, Optional arguments) As String

    Dim msg_list As String
    Dim err_msg As String
    Dim msg_offset As Long
    Dim msg_start As Long
    Dim msg_end As Long
    Dim argument_index As Integer
    
    msg_list = ERR_MSG_LIST & "|"
    msg_offset = InStr(msg_list, "|#" & err_id & "|")
    If msg_offset = 0 Then
        get_err_msg = "error " & err_id & " detected, no err msg available"
        Exit Function
    End If
    msg_start = InStr(msg_offset + 1, msg_list, "|") + 1
    msg_end = InStr(msg_start + 1, msg_list, "|") - 1
    err_msg = Mid(msg_list, msg_start, msg_end - msg_start + 1)
    err_msg = Replace(err_msg, "#{pipe}", "|")
    If IsArray(arguments) Then
        For argument_index = 0 To UBound(arguments)
            err_msg = Replace(err_msg, "#{" & argument_index + 1 & "}", arguments(argument_index))
        Next
    End If
    get_err_msg = err_msg
End Function

Public Sub raise(err_id As Long, function_name As String, Optional description)
    LastFailMsg = description
    Err.raise err_id, function_name, description
End Sub

Public Property Get LastFailMsg() As String
    
    LastFailMsg = m_last_fail_msg
End Property

Public Property Let LastFailMsg(ByVal last_fail_msg As String)
    ' this is a workaround for VBAs mishandling of custom errors
    '  e.g. using callbyname() will overwritecustom errors with 440 automation error
    '   https://stackoverflow.com/questions/18241906/vba-callbyname-and-invokehook
    '  e.g. using strongly typed objects will overwrite the error description
    '   https://stackoverflow.com/questions/31234805/err-raise-is-ignoring-custom-description-and-source
    ' so in addition of adding expectation results to custom errors (aka exceptions), the most current result meessage will be saved here

    m_last_fail_msg = last_fail_msg
End Property
