Attribute VB_Name = "TFeature"
Option Explicit

Public Function fail_step(err_id As Long, Optional err_msg) As Variant
 
    Dim err_desc As String
    
    If IsMissing(err_msg) Then
        err_desc = vbNullString
    Else
        err_desc = err_msg
    End If
    Select Case err_id
    Case ERR_ID_STEP_IS_PENDING
        fail_step = Array("PENDING", err_desc)
    Case ERR_ID_STEP_IS_MISSING
        fail_step = Array("MISSING", "PENDING: add code snippet as suggestion for missing test steps")
    Case Else
        fail_step = Array("FAILED", err_desc)
    End Select
End Function

Public Sub pending(pending_msg)
    Err.Raise ERR_ID_STEP_IS_PENDING, Description:=pending_msg
End Sub

