Attribute VB_Name = "TCase"
Option Explicit

Dim m_Logger As logger

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

Public Property Get Log() As logger
    
    If m_Logger Is Nothing Then
        Set m_Logger = New logger
    End If
    Set Log = m_Logger
End Property

Public Property Let Log(ByVal new_Logger As logger)
    Set m_Logger = new_Logger
End Property
