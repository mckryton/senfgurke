Attribute VB_Name = "TFeature"
Option Explicit

Dim m_logger As Logger

Public Function raise_step_error(err_id As Long, Optional err_msg) As Variant
 
    Dim err_desc As String
    
    If IsMissing(err_msg) Then
        err_desc = ""
    Else
        err_desc = err_msg
    End If
    TExampleRunner.stop_test
    Select Case err_id
    Case ERR_ID_STEP_IS_PENDING
        raise_step_error = Array("PENDING", err_desc)
    Case ERR_ID_STEP_IS_MISSING
        raise_step_error = Array("MISSING")
    Case Else
        raise_step_error = Array("FAILED", err_desc)
    End Select
End Function

Public Sub pending(pending_msg)
    Err.raise ERR_ID_STEP_IS_PENDING, description:=pending_msg
End Sub

Public Property Get Log() As Logger
    
    If m_logger Is Nothing Then
        Set m_logger = New Logger
    End If
    Set Log = m_logger
End Property

Public Property Let Log(ByVal new_Logger As Logger)
    Set m_logger = new_Logger
End Property
